The workflow is:

Start with the NSFR_Template.xlsx file (which contains your dashboards, charts, and formulas showing #REF! errors).

The Java program will open this template in memory.

It will create the three data sheets (DWH Dates, DWH Day, LKP_IFRSLUX) and populate them with fresh data from the SQL queries.

It will then force Excel's calculation engine to update all formulas across the entire workbook. The formulas on your dashboard will now find the data in the new sheets and calculate correctly.

Finally, it will save the result as a new, timestamped, read-only report file.

I have updated the Java code to follow this exact logic.

Part 1: Prepare Your Template
Before running the code, make sure you have the NSFR_Template.xlsx file ready, as we discussed in the previous step:

It's a copy of your original file.

The data sheets (DWH Dates, DWH Day, etc.) and other unneeded sheets (CAPITAUX PROPRES ET RECAP, etc.) have been deleted.

The broken Named Ranges have been cleaned up via the Name Manager.

Part 2: The Updated Java Code
This new class, ReportGeneratorFromTemplate, is designed to perform the complete workflow. It now loads the template, handles existing sheets, and forces the critical formula recalculation at the end.

Java

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ReportGeneratorFromTemplate {

    // --- CONFIGURATION ---
    private static final String TEMPLATE_FILE_PATH = "NSFR_Template.xlsx"; // The path to your template file
    private static final String OUTPUT_FILE_PATH = "NSFR_Report_" + new SimpleDateFormat("yyyy-MM-dd").format(new Date()) + ".xlsx";

    // This connection string is built based on your report (Windows Authentication).
    // You might need to add 'encrypt=true;trustServerCertificate=true;' for modern drivers.
    private static final String DB_CONNECTION_STRING = "jdbc:sqlserver://SRVDWH;databaseName=NCIBL;integratedSecurity=true;encrypt=true;trustServerCertificate=true;";

    // !!! IMPORTANT: REPLACE THESE WITH YOUR ACTUAL SQL QUERIES !!!
    private static final String SQL_FOR_DWH_DATES = "SELECT * FROM YourTableForDwhDates;"; // Replace this
    private static final String SQL_FOR_DWH_DAY = "SELECT * FROM YourTableForDwhDay;";   // Replace this
    private static final String SQL_FOR_LKP_IFRSLUX = "SELECT * FROM YourTableForIfrsLux;"; // Replace this


    public static void main(String[] args) {
        System.out.println("Starting report generation process from template...");

        try (FileInputStream templateFileStream = new FileInputStream(TEMPLATE_FILE_PATH);
             XSSFWorkbook workbook = new XSSFWorkbook(templateFileStream)) {

            System.out.println("Template file '" + TEMPLATE_FILE_PATH + "' loaded successfully.");

            // Step 1: Inject the three data sheets into the workbook from the database
            generateSheetFromQuery(workbook, "DWH Dates", SQL_FOR_DWH_DATES);
            generateSheetFromQuery(workbook, "DWH Day", SQL_FOR_DWH_DAY);
            generateSheetFromQuery(workbook, "LKP_IFRSLUX", SQL_FOR_LKP_IFRSLUX);

            // Step 2: Force recalculation of all formulas in the entire workbook
            recalculateAllFormulas(workbook);

            // Step 3: Write the final, calculated workbook to a new file
            System.out.println("Saving final report to '" + OUTPUT_FILE_PATH + "'...");
            try (FileOutputStream outputStream = new FileOutputStream(OUTPUT_FILE_PATH)) {
                workbook.write(outputStream);
            }
            
            // Step 4: Set the final report to be read-only
            File finalReport = new File(OUTPUT_FILE_PATH);
            if (finalReport.exists()) {
                finalReport.setReadOnly();
                System.out.println("Final report has been set to read-only.");
            }
            
            System.out.println("Process completed successfully.");

        } catch (IOException | SQLException e) {
            System.err.println("A critical error occurred during the report generation.");
            e.printStackTrace();
        }
    }

    /**
     * Runs a query and writes the results to a sheet. If the sheet already exists, it is deleted and recreated.
     */
    private static void generateSheetFromQuery(XSSFWorkbook workbook, String sheetName, String sqlQuery) throws SQLException {
        System.out.println("  -> Generating data for sheet: '" + sheetName + "'...");

        // Remove the sheet if it already exists to ensure a clean slate
        int existingSheetIndex = workbook.getSheetIndex(sheetName);
        if (existingSheetIndex != -1) {
            workbook.removeSheetAt(existingSheetIndex);
        }
        XSSFSheet sheet = workbook.createSheet(sheetName);

        try (Connection conn = DriverManager.getConnection(DB_CONNECTION_STRING);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sqlQuery)) {

            ResultSetMetaData metaData = rs.getMetaData();
            int columnCount = metaData.getColumnCount();

            // Create Header Row
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = createHeaderStyle(workbook);
            for (int i = 1; i <= columnCount; i++) {
                Cell cell = headerRow.createCell(i - 1);
                cell.setCellValue(metaData.getColumnLabel(i));
                cell.setCellStyle(headerStyle);
            }

            // Write Data Rows
            int rowNum = 1;
            while (rs.next()) {
                Row row = sheet.createRow(rowNum++);
                for (int i = 1; i <= columnCount; i++) {
                    Cell cell = row.createCell(i - 1);
                    Object value = rs.getObject(i);
                    // Set cell value based on its type
                    if (value instanceof String) {
                        cell.setCellValue((String) value);
                    } else if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else if (value instanceof java.sql.Date || value instanceof java.sql.Timestamp) {
                        CellStyle dateStyle = workbook.createCellStyle();
                        CreationHelper createHelper = workbook.getCreationHelper();
                        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
                        cell.setCellValue((Date) value);
                        cell.setCellStyle(dateStyle);
                    } else if (value != null) {
                        cell.setCellValue(value.toString());
                    }
                }
            }
            
            for(int i = 0; i < columnCount; i++) {
                sheet.autoSizeColumn(i);
            }
            System.out.println("     Done. Wrote " + (rowNum - 1) + " rows of data to '" + sheetName + "'.");
        }
    }

    /**
     * CRITICAL STEP: This method iterates through all formulas in the workbook and calculates their results.
     */
    private static void recalculateAllFormulas(XSSFWorkbook workbook) {
        System.out.println("Recalculating all formulas in the workbook...");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        evaluator.evaluateAll();
        System.out.println(" -> Formula recalculation complete.");
    }

    /**
     * Helper method to create a bold style for header cells.
     */
    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
}
Part 3: Instructions for Use
Place the Template: Put your cleaned NSFR_Template.xlsx file in the root directory of your Java project, or update the TEMPLATE_FILE_PATH constant with the correct path.

Update the Code: Create a new Java class ReportGeneratorFromTemplate.java, paste the code above, and replace the placeholder SQL queries with your real ones.

Run the main method.

When you run it, you will see a console output showing the steps. The final result will be a new file, for example, NSFR_Report_2025-10-06.xlsx. When you open this new file, the Dashboard_SdM and other sheets should be fully calculated and showing the correct values, just like your original file did.
