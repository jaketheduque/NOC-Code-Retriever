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
import org.openqa.selenium.htmlunit.HtmlUnitDriver;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.ResourceBundle;

public class NVCourtWebscraper {
    private static final Logger log = LoggerFactory.getLogger(NVCourtWebscraper.class);

    private static final ResourceBundle PROPERTIES = ResourceBundle.getBundle("application");
    private static final String COURT_URL = PROPERTIES.getString("court_url");

    public static void main(String[] args) throws ClassNotFoundException {
        log.info("Launching headless browser");

        // Instantiates a new headless WebDriver object
        WebDriver driver = new HtmlUnitDriver();

        driver.get(COURT_URL);

        log.info("Court website loaded");

        WebElement element = driver.findElement(By.xpath("//a[contains(text(), '(3) NOC_NIBRS (current)')]"));
        String downloadURL = element.getAttribute("href");

        log.info("Download URL obtained");
        log.info("URL: " + downloadURL);
        log.info("Closing headless browser");

        driver.close();

        // Saves the file to temp.xls
        try (BufferedInputStream in = new BufferedInputStream(new URL(downloadURL).openStream());
             FileOutputStream out = new FileOutputStream("temp.xls")) {
            byte dataBuffer[] = new byte[1024];
            int bytesRead;
            while ((bytesRead = in.read(dataBuffer, 0, 1024)) != -1) {
                out.write(dataBuffer, 0, bytesRead);
            }
        } catch (IOException e) {
            log.error("Error saving file! Please check exception");
            e.printStackTrace();
        }

        log.info("Temporary file saved");

        // Parses the saved xls file and prepares an array for population
        NOCEntry[] entries = null;
        try {
            Workbook workbook = new HSSFWorkbook(new FileInputStream("temp.xls"));

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
            new File("temp.xls").delete();
        }

        /** Implement SQL at later date
        // Loads the SQL driver
        Class.forName("oracle.jdbc.driver.OracleDriver");
        String url = PROPERTIES.getString("sql_url");
        String username = PROPERTIES.getString("sql_username");
        String password = PROPERTIES.getString("sql_password");
        */
    }
}
