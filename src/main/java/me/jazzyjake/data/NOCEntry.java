package me.jazzyjake.data;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class NOCEntry {
    private final int noc;
    private final String degree;
    private final String description;
    private final String reportable;

    public NOCEntry (Row r) {
        Cell noc = r.getCell(1);
        Cell degree = r.getCell(2);
        Cell description = r.getCell(6);
        Cell reportable = r.getCell(14);

        String reportableString = reportable != null ? reportable.getStringCellValue() : "N";

        this.noc = (int) noc.getNumericCellValue();
        this.degree = degree.getStringCellValue();
        this.description = description.getStringCellValue();
        this.reportable = reportableString;
    }

    public NOCEntry() {
        this.noc = 0;
        this.degree = null;
        this.description = null;
        this.reportable = null;
    }

    public int getNoc() {
        return noc;
    }

    public String getDegree() {
        return degree;
    }

    public String getDescription() {
        return description;
    }

    public String getReportable() {
        return reportable;
    }

    @Override
    public String toString() {
        return String.format("NOC: %d, Degree: %s, Description: %s, Reportable: %s", noc, degree, description, reportable);
    }
}
