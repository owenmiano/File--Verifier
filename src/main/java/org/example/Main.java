package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import java.text.SimpleDateFormat;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Main {

    private static final String CSV_FILE_NAME = "member_details.csv";
    private static final Path csvFilePath = Paths.get("src", "main", "resources", CSV_FILE_NAME);
    private static final Pattern ID_PATTERN = Pattern.compile("\\d{8}");
    private static final Pattern MOBILE_PATTERN = Pattern.compile("\\d{10}");
    private static final Pattern EMAIL_PATTERN = Pattern.compile("^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$");
    private static final String[] HEADER_COLUMNS = {"ID Number", "Name", "Phone Number", "Email", "Gender"};
    private static final String[] HEADER_COLUMNS_WITH_ERRORS = {"ID Number", "Name", "Phone Number", "Email", "Gender", "Errors"};

    public static void main(String[] args) {
        System.out.println("Sanitizing your data.Sit tight");
        long programStartTime = System.currentTimeMillis();

        try {
            processCSVAndExportToExcel();
            System.out.println("Data Sanitization completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        long programEndTime = System.currentTimeMillis();
        double programTimeMinutes = (programEndTime - programStartTime) / 60000.0;
        System.out.println("Total time taken to process data " + programTimeMinutes + " minutes");
    }

    private static void processCSVAndExportToExcel() throws IOException {
        // Append timestamp to the filename
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy_MM_dd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        String newExcelFileName = "src/main/resources/" + timestamp+ "_sanitized" + ".xlsx"; // Include the path
        Workbook workbook = new SXSSFWorkbook();

        Sheet maleSheet = workbook.createSheet("Male");
        createHeaderRow(maleSheet, HEADER_COLUMNS);

        Sheet femaleSheet = workbook.createSheet("Female");
        createHeaderRow(femaleSheet, HEADER_COLUMNS);

        Sheet invalidSheet = workbook.createSheet("Invalid Records");
        createHeaderRow(invalidSheet, HEADER_COLUMNS_WITH_ERRORS);

        try (Stream<String> stream = Files.lines(csvFilePath)) {
            List<String[]> rows = stream.skip(1) // Skip header row
                    .map(line -> line.split(","))
                    .collect(Collectors.toList());

            for (String[] columns : rows) {
                processRow(columns, maleSheet, femaleSheet, invalidSheet);
            }
        }


        try (FileOutputStream outputStream = new FileOutputStream(newExcelFileName)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    private static void processRow(String[] columns, Sheet maleSheet, Sheet femaleSheet, Sheet invalidSheet) {
        String idNumber = columns[0].trim();
        String mobileNumber = columns[2].trim();
        String email = columns[3].trim();
        String gender = columns[4].trim();

        StringBuilder errors = new StringBuilder();

        if (!isValid(idNumber, ID_PATTERN)) {
            errors.append("Invalid ID Number, ");
        }

        if (!isValid(mobileNumber, MOBILE_PATTERN)) {
            errors.append("Invalid Phone Number, ");
        }

        if (!isValidEmail(email)) {
            errors.append("Invalid Email, ");
        }

        if (!errors.isEmpty()) {
            errors.setLength(errors.length() - 2);
            if (columns.length < 6) {
                columns = Arrays.copyOf(columns, 6);
            }
            columns[5] = errors.toString();
            createDataRow(invalidSheet, columns);
        } else {
            Sheet targetSheet = getTargetSheet(gender, maleSheet, femaleSheet, invalidSheet);
            createDataRow(targetSheet, columns);
        }
    }

    private static boolean isValid(String value, Pattern pattern) {
        return pattern.matcher(value).matches();
    }

    private static boolean isValidEmail(String value) {
        return EMAIL_PATTERN.matcher(value).matches();
    }

    private static Sheet getTargetSheet(String gender, Sheet maleSheet, Sheet femaleSheet, Sheet invalidSheet) {
        if ("Male".equalsIgnoreCase(gender)) {
            return maleSheet;
        } else if ("Female".equalsIgnoreCase(gender)) {
            return femaleSheet;
        } else {
            return invalidSheet;
        }
    }

    private static void createHeaderRow(Sheet sheet, String[] headers) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
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