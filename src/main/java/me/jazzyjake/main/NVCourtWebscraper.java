package me.jazzyjake.main;

import me.jazzyjake.data.NOCEntry;
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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.ResourceBundle;
import java.util.stream.Stream;

public class NVCourtWebscraper {
    private static final Logger log = LoggerFactory.getLogger(NVCourtWebscraper.class);

    private static final ResourceBundle PROPERTIES = ResourceBundle.getBundle("application");
    private static final String COURT_URL = PROPERTIES.getString("court_url");
    private static final String DOWNLOAD_FILEPATH = PROPERTIES.getString("temp_file_download_path");
    private static final String CHROMEDRIVER_PATH = PROPERTIES.getString("chromedriver_path");

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
            e.printStackTrace();
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
            log.error("Error working with Excel file! Please check exception");
            e.printStackTrace();
        } finally {
            log.info("Deleting temporary file");

            // Deletes the temp.xml file
            new File(filename).delete();

            log.info("Terminating all chromedriver processes");
            Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe /T");
        }

        // Parameters to connect to database
        String url = PROPERTIES.getString("db_url");
        String name = PROPERTIES.getString("db_name");
        String username = PROPERTIES.getString("db_username");
        String password = PROPERTIES.getString("db_password");

        String connectionUrl =
                "jdbc:sqlserver://" + url + ":1433;"
                        + "database=" + name + ";"
                        + "user=" + username + ";"
                        + "password=" + password + ";"
                        + "encrypt=true;trustServerCertificate=true;loginTimeout=30;";

        log.info("Connecting to NOC database");
        // Connects to production database
        try (Connection connection = DriverManager.getConnection(connectionUrl)) {
            // Clears table contents before inserting new data
            PreparedStatement stmt = connection.prepareStatement("DELETE FROM dbo.NOCList");
            stmt.execute();

            log.info("Previous table entries cleared");

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
                } catch (SQLException ex) {
                    ex.printStackTrace();
                }
            }

            log.info("About to add {} entries to NOC database", (id - 1));

            // Executes the SQL update
            stmt.clearParameters();
            int[] results = stmt.executeBatch();

            log.info("Added entries", results.length);
        }
        // Handle any errors that may have occurred.
        catch (SQLException e) {
            e.printStackTrace();
        }
    }
}
