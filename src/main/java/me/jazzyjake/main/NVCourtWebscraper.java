package me.jazzyjake.main;

import me.jazzyjake.data.NOCEntry;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.*;
import java.util.*;
import java.util.Date;
import java.util.stream.Stream;

public class NVCourtWebscraper {
    private static final Logger log = LogManager.getLogger(NVCourtWebscraper.class);

    private static final ResourceBundle PROPERTIES = ResourceBundle.getBundle("application");
    private static final String COURT_URL = PROPERTIES.getString("court_url");
    private static final String DOWNLOAD_FILEPATH = PROPERTIES.getString("temp_file_download_path");
    private static final String CHROMEDRIVER_PATH = PROPERTIES.getString("chromedriver_path");

    private static final String REPORTING_DB_NAME = PROPERTIES.getString("reporting_db_name");
    private static final String COURTVIEW_DB_NAME = PROPERTIES.getString("courtview_db_name");
    private static final String EXCHANGE_DB_NAME = PROPERTIES.getString("exchange_db_name");

    private static final String DEV_EMAIL = PROPERTIES.getString("dev_email_address");
    private static final String NOTIFICATION_EMAIL = PROPERTIES.getString("notification_email_address");

    public static void main(String[] args) throws InterruptedException, IOException {
        log.info("Launching Selenium browser");

        System.setProperty("webdriver.chrome.driver", CHROMEDRIVER_PATH);

        // Sets up ChromeDriver download folder location
        HashMap<String, Object> chromePrefs = new HashMap<>();
        chromePrefs.put("profile.default_content_settings.popups", 0);
        chromePrefs.put("download.default_directory", DOWNLOAD_FILEPATH);
        ChromeOptions options = new ChromeOptions();
        options.setExperimentalOption("prefs", chromePrefs);

        // Instantiates a new headless WebDriver object
        WebDriver driver = new ChromeDriver(options);

        driver.get(COURT_URL);

        log.info("Court website loaded");

        WebElement element = driver.findElement(By.xpath("//a[contains(text(), '(3) NOC_NIBRS (current)')]"));

        // Clicks the file, thus downloading the file
        element.click();

        log.info("Downloading file!");

        Thread.sleep(20 * 1000);

        log.info("Closing Chrome browser");

        driver.quit();

        log.info("Temporary file saved");

        String filename = null;
        // Gets (what should be) the only xls file in the temp directory
        try (Stream<Path> walk = Files.walk(Paths.get(DOWNLOAD_FILEPATH))) {
            filename = walk
                    .filter(p -> !Files.isDirectory(p))
                    .map(p -> p.toString().toLowerCase())
                    .filter(f -> f.endsWith("xls"))
                    .findFirst()
                    .get();
        } catch (IOException e) {
            log.error("Error occurred while downloading NOC file from court website! Please check exception", e);
            emailLogFile(DEV_EMAIL, "Error occurred while downloading NOC file from court website");
        }

        log.info("Downloaded File Filename: " + filename);

        // Parses the saved xls file and prepares an array for population
        NOCEntry[] entries = null;
        try {
            Workbook workbook = new HSSFWorkbook(new FileInputStream(filename));

            Sheet sheet = workbook.getSheetAt(0);

            // Creates an ArrayList for the NOCEntry objects
            List<NOCEntry> entriesList = new ArrayList<>();
            int i = 0;
            for (Row r : sheet) {
                // Skip first row
                if (r.getRowNum() == 0) continue;

                Cell noc = r.getCell(1);
                Cell degree = r.getCell(2);
                Cell description = r.getCell(6);

                // Makes sure the row is valid
                if (noc != null && degree != null && description != null) {
                    NOCEntry entry = new NOCEntry(r);
                    entriesList.add(entry);
                }
            }

            // Converts the ArrayList to an array
            entries = entriesList.toArray(new NOCEntry[0]);
        } catch (IOException e) {
            log.error("Error working with Excel file! Please check exception", e);
            emailLogFile(DEV_EMAIL, "Error working with Excel file! Please check exception");
        } finally {
            log.info("Deleting temporary file");

            // Deletes the temp.xml file
            new File(filename).delete();

            log.info("Terminating all chromedriver processes");
            Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe /T");
        }

        // Parameters to connect to Reporting database
        String url = PROPERTIES.getString("db_url");
        String username = PROPERTIES.getString("db_username");
        String password = PROPERTIES.getString("db_password");

        String connectionUrl =
                "jdbc:sqlserver://" + url + ":1433;"
                        + "database=" + REPORTING_DB_NAME + ";"
                        + "user=" + username + ";"
                        + "password=" + password + ";"
                        + "encrypt=true;trustServerCertificate=true;loginTimeout=30;";

        log.info("Connecting to NOC Reporting database");
        // Connects to production database
        try (Connection connection = DriverManager.getConnection(connectionUrl)) {
            // Clears table contents before inserting new data
            PreparedStatement stmt = connection.prepareStatement("DELETE FROM dbo.NOCList");
            stmt.execute();

            log.info("Previous Reporting database table entries cleared");

            // Prepares INSERT statement
            stmt = connection.prepareStatement("INSERT INTO dbo.NOCList VALUES (?, ?, ?, ?, ?)");

            // Adds each row of the xls file to the SQL statement
            int id = 1;
            for (NOCEntry e : entries) {
                try {
                    stmt.clearParameters();
                    stmt.setInt(1, id++);
                    stmt.setInt(2, e.getNoc());
                    stmt.setString(3, e.getDescription());
                    stmt.setString(4, e.getReportable());
                    stmt.setString(5, e.getDegree());

                    stmt.addBatch();
                } catch (SQLException throwables) {
                    log.error("SQL Exception occurred! Please check exception", throwables);
                    emailLogFile(DEV_EMAIL, "SQL Exception occurred! Please check exception");
                }
            }

            log.info("About to add {} entries to NOC Reporting database", (id - 1));

            // Executes the SQL update
            stmt.clearParameters();
            int[] results = stmt.executeBatch();

            log.info("Added {} entries to Reporting database", results.length);

            log.info("Now moving on to updating {} and {} databases", EXCHANGE_DB_NAME, COURTVIEW_DB_NAME);

            // Updates Courtview database
            updateWithNewNOCCodeValues(connection, COURTVIEW_DB_NAME, "N", "OffenseCode");

            // Updates exchange database
            updateWithNewNOCCodeValues(connection, EXCHANGE_DB_NAME, "N", "OffenseCode");
            updateWithNewNOCCodeValues(connection, EXCHANGE_DB_NAME, "Y", "ReportableOffenseCode");
        }
        // Handle any errors that may have occurred.
        catch (SQLException e) {
            log.error("SQL Exception occurred! Please check exception", e);
            emailLogFile(DEV_EMAIL, "SQL Exception occurred! Please check exception");
        }

        log.info("Broker NOC update complete!");

        emailLogFile(NOTIFICATION_EMAIL, "Broker NOC Update Completed Successfully!");
    }

    /**
     * Updates provided database with new NOC records from Reporting database using SQL stored procedures
     *
     * @param conn SQL connection to use
     * @param updateDB Database to be updated with new NOC codes
     * @param reportable Reportable (Y/N)
     * @param type Type of code to be inserted (ex. OffenseCode)
     */
    private static void updateWithNewNOCCodeValues(Connection conn, String updateDB, String reportable, String type) throws SQLException {
        // Gets new codes from Reporting database
        PreparedStatement stmt = conn.prepareStatement("{call sp_SYS_GetNewNOCS(?, ?, ?)}");
        stmt.setString(1, updateDB);
        stmt.setString(2, reportable);
        stmt.setString(3, REPORTING_DB_NAME);
        ResultSet rs = stmt.executeQuery();

        // Prepares SQL stored procedure statement to update specified database row by row with new NOC entries
        stmt = conn.prepareStatement("{call sp_SYS_InsertCodeValue(?, ?, ?)}");

        // Adds each new NOC code from Reporting database to batch SQL statement
        int count = 0;
        while (rs.next()) {
            stmt.clearParameters();
            stmt.setInt(1, rs.getInt(2));
            stmt.setString(2, rs.getString(3));
            stmt.setString(3, type);
            stmt.addBatch();
            count++;
        }

        log.info("About to add {} new NOC entries from Reporting database to {} database", count, updateDB);

        // Runs update batch
        int[] results = stmt.executeBatch();

        log.info("Added {} new NOC entries to {} database", results.length, updateDB);
    }

    /**
     * Sends the currently saved log file to the email provided using the specified email server in properties file
     *
     * @param email Email address to send log file to
     * @param subject Subject field of the email
     */
    private static void emailLogFile(String email, String subject) {
        try {
            // Puts the SMTP server ip into system properties
            Properties props = System.getProperties();
            props.put("mail.smtp.host", PROPERTIES.getString("email_server_ip"));

            // Opens email server session
            Session session = Session.getInstance(props);

            // Creates new message and sets headers
            MimeMessage msg = new MimeMessage(session);
            msg.addHeader("Content-type", "text/HTML; charset=UTF-8");
            msg.addHeader("format", "flowed");
            msg.addHeader("Content-Transfer-Encoding", "8bit");

            // Sets the email address the email is from and to reply to
            msg.setFrom(new InternetAddress(PROPERTIES.getString("replay_email_address")));
            msg.setReplyTo(InternetAddress.parse(PROPERTIES.getString("replay_email_address"), false));

            // Sets subject and send datetime of email
            msg.setSubject(subject, "UTF-8");
            msg.setSentDate(new Date());

            // Attaches the log file to the email using Multipart email body
            Multipart multipart = new MimeMultipart();
            BodyPart body = new MimeBodyPart();
            DataSource source = new FileDataSource("logs/logfile.log");
            body.setDataHandler(new DataHandler(source));
            body.setFileName("logs/logfile.log");
            multipart.addBodyPart(body);
            msg.setContent(multipart);

            msg.setRecipients(Message.RecipientType.TO, InternetAddress.parse(email, false));
            Transport.send(msg);
        } catch (MessagingException e) {
            log.error("Error occurred sending log file email!", e);
        }
    }
}
