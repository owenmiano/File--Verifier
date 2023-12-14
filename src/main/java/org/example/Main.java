package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.util.regex.Pattern;

public class Main {

    //Path to CSV and Excel
    private static final String CSV_FILE_PATH = "C:\\Users\\Home\\Downloads\\member_details.csv";
    private static final String EXCEL_FILE_PATH = "C:\\Users\\Home\\Downloads\\sanitized.xlsx";

    private static final String[] HEADER_COLUMNS = {"ID Number", "Name", "Phone Number", "Email", "Gender"};

    public static void main(String[] args) {
        long programStartTime = System.currentTimeMillis();
        try {
            processCSVAndExportToExcel();
            System.out.println("Processing completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
        long programEndTime = System.currentTimeMillis();
        long programTimeMillis = programEndTime - programStartTime;
        double programTimeMinutes = programTimeMillis / 60000.0;
        System.out.println("Total time taken to process data " + programTimeMinutes + " minutes");
    }

    private static void processCSVAndExportToExcel() throws IOException {
        try (BufferedReader reader = new BufferedReader(new FileReader(CSV_FILE_PATH))) {
            Workbook workbook = new SXSSFWorkbook();

            Sheet maleSheet = workbook.createSheet("Male");
            createHeaderRow(maleSheet, HEADER_COLUMNS);

            Sheet femaleSheet = workbook.createSheet("Female");
            createHeaderRow(femaleSheet, HEADER_COLUMNS);

            Sheet invalidSheet = workbook.createSheet("Invalid Records");
            createHeaderRow(invalidSheet, HEADER_COLUMNS);

            String line;
            int count = 1;
            System.out.println("Cleaning your data sit tight");
            while ((line = reader.readLine()) != null) {
                String[] columns = line.split(",");

                if (columns.length == 5) {
                    String idNumber = columns[0].trim();
                    String mobileNumber = columns[2].trim();
                    String email = columns[3].trim();
                    String gender = columns[4].trim();

                    if (isValid(idNumber) && isValidMobileNumber(mobileNumber) && isValidEmail(email)) {
                        Sheet targetSheet;
                        if ("Male".equalsIgnoreCase(gender)) {
                            targetSheet = maleSheet;
                        } else if ("Female".equalsIgnoreCase(gender)) {
                            targetSheet = femaleSheet;
                        } else {
                            targetSheet = invalidSheet;
                        }
                        createDataRow(targetSheet, columns);
                    } else {
                        createDataRow(invalidSheet, columns);
                    }
                }
                System.out.println("counting " + count);
                count++;
            }

            try (FileOutputStream outputStream = new FileOutputStream(EXCEL_FILE_PATH)) {
                workbook.write(outputStream);
            }
        }
    }

    private static void createHeaderRow(Sheet sheet, String[] headers) {
        // Check if the sheet already has a header row
        if (sheet.getPhysicalNumberOfRows() == 0) {
            Row headerRow = sheet.createRow(0);

            for (int i = 0; i < headers.length; i++) {
                headerRow.createCell(i).setCellValue(headers[i]);
            }
        }
    }

    private static boolean isValid(String value) {
        // Validate ID Number
        return value != null && value.trim().matches("\\d{8}");
    }

    private static boolean isValidMobileNumber(String value) {
        // Validate mobile number
        return value != null && value.trim().matches("\\d{10}");
    }

    private static boolean isValidEmail(String value) {
        // Validate email address
        return Pattern.compile("^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$").matcher(value).matches();
    }

    private static void createDataRow(Sheet sheet, String[] data) {
        int rowCount = sheet.getPhysicalNumberOfRows();

        // Check if the row count is within the allowable range
        if (rowCount >= 1048575) {
            return;
        }

        Row row = sheet.createRow(rowCount);

        for (int i = 0; i < data.length; i++) {
            row.createCell(i).setCellValue(data[i]);
        }
    }
}
