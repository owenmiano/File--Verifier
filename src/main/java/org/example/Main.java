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
import java.util.concurrent.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class Main {
    private static final String CSV_FILE_NAME = "member_details.csv";
    private static final Path csvFilePath = Paths.get("src", "main", "resources", CSV_FILE_NAME);
    private static final Pattern ID_PATTERN = Pattern.compile("\\d{8}");
    private static final Pattern MOBILE_PATTERN = Pattern.compile("\\d{10}");
    private static final Pattern EMAIL_PATTERN = Pattern.compile("^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$");
    private static final String[] HEADER_COLUMNS = {"ID Number", "Name", "Phone Number", "Email", "Gender"};
    private static final String[] HEADER_COLUMNS_WITH_ERRORS = {"ID Number", "Name", "Phone Number", "Email", "Gender", "Errors"};
    private static final int BATCH_SIZE = 1000;
    private static final int THREAD_COUNT = 4;

    public static void main(String[] args) {
        System.out.println("Sanitizing your data. Sit tight");
        long programStartTime = System.currentTimeMillis();

        try {
            processCSVAndExportToExcel();
            System.out.println("Data Sanitization completed successfully.");
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        } catch (ExecutionException e) {
            throw new RuntimeException(e);
        }

        long programEndTime = System.currentTimeMillis();
        double programTimeMinutes = (programEndTime - programStartTime) / 60000.0;
        System.out.println("Total time taken to process data " + programTimeMinutes + " minutes");
    }

    private static void processCSVAndExportToExcel() throws IOException, InterruptedException, ExecutionException {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy_MM_dd_HHmmss");
        String timestamp = dateFormat.format(new Date());
        String newExcelFileName = "src/main/resources/" + timestamp + "_sanitized.xlsx";
        Workbook workbook = new SXSSFWorkbook();

        Sheet maleSheet = workbook.createSheet("Male");
        createHeaderRow(maleSheet, HEADER_COLUMNS);

        Sheet femaleSheet = workbook.createSheet("Female");
        createHeaderRow(femaleSheet, HEADER_COLUMNS);

        Sheet invalidSheet = workbook.createSheet("Invalid Records");
        createHeaderRow(invalidSheet, HEADER_COLUMNS_WITH_ERRORS);

        List<String[]> rows;
        try (Stream<String> stream = Files.lines(csvFilePath)) {
            rows = stream.skip(1)
                    .map(line -> line.split(","))
                    .collect(Collectors.toList());
        }

        ExecutorService executorService = Executors.newFixedThreadPool(THREAD_COUNT);
        List<Future<List<String[]>>> futures = new ArrayList<>();

        for (int i = 0; i < rows.size(); i += BATCH_SIZE) {
            int start = i;
            int end = Math.min(i + BATCH_SIZE, rows.size());
            List<String[]> batch = rows.subList(start, end);
            futures.add(executorService.submit(() -> processBatch(batch)));
        }

        for (Future<List<String[]>> future : futures) {
            List<String[]> processedRows = future.get();
            for (String[] columns : processedRows) {
                Sheet targetSheet = getTargetSheet(columns, maleSheet, femaleSheet, invalidSheet);
                createDataRow(targetSheet, columns);
            }
        }

        executorService.shutdown();
        executorService.awaitTermination(1, TimeUnit.MINUTES);

        try (FileOutputStream outputStream = new FileOutputStream(newExcelFileName)) {
            workbook.write(outputStream);
        }
    }

    private static List<String[]> processBatch(List<String[]> batch) {
        return batch.stream()
                .map(Main::validateAndProcessRow)
                .collect(Collectors.toList());
    }

    private static String[] validateAndProcessRow(String[] columns) {
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

        if (errors.length() > 0) {
            errors.setLength(errors.length() - 2);
            if (columns.length < 6) {
                columns = Arrays.copyOf(columns, 6);
            }
            columns[5] = errors.toString();
        }
        return columns;
    }

    private static boolean isValid(String value, Pattern pattern) {
        return pattern.matcher(value).matches();
    }

    private static boolean isValidEmail(String value) {
        return EMAIL_PATTERN.matcher(value).matches();
    }

    private static Sheet getTargetSheet(String[] columns, Sheet maleSheet, Sheet femaleSheet, Sheet invalidSheet) {
        if (columns.length > 5 && columns[5] != null) {
            return invalidSheet;
        }
        String gender = columns[4].trim();
        if ("Male".equalsIgnoreCase(gender)) {
            return maleSheet;
        } else if ("Female".equalsIgnoreCase(gender)) {
            return femaleSheet;
        }
        return invalidSheet;
    }

    private static void createHeaderRow(Sheet sheet, String[] headers) {
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    private static void createDataRow(Sheet sheet, String[] data) {
        int rowCount = sheet.getPhysicalNumberOfRows();
        if (rowCount >= 1048575) {
            return;
        }
        Row row = sheet.createRow(rowCount);
        for (int i = 0; i < data.length; i++) {
            row.createCell(i).setCellValue(data[i]);
        }
    }
}
