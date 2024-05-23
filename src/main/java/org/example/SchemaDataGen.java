package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
public class SchemaDataGen {
    private static final Set<String> filenameSet = new HashSet<>();
    private static final Map<String, Integer> headerMap = new HashMap<>();
    private static String templateFilePath;
    private static String excelFilePath;
    private static final String outputRootFolder = System.getProperty("user.dir");

    public static void main(String[] args) {
        try {
            if (args.length < 2) {
                System.out.println("Usage: java -jar YourJarName.jar <ExcelFilePath> <TemplateFilePath>");
                System.exit(1);
            }
            excelFilePath = args[0];
            templateFilePath = args[1];
            processExcelFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void processExcelFile() throws IOException {
        try (InputStream excelInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(excelInputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            processTemplate(sheet.getRow(0)); // Process the header row to extract placeholders
            String outputFolderPath = createOutputFolder();
            AtomicInteger rowNumber = new AtomicInteger(1);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    // Skip header row
                    continue;
                }
                processRow(row, outputFolderPath, rowNumber);
            }
        }
    }

    private static void processRow(Row row, String outputFolderPath, AtomicInteger rowNumber) {
        // Extract the first column as the file name
        String fileName = cellToString(row.getCell(0));

        // Validate filename to ensure no duplicates
        if (!isValidFilename(fileName)) {
            System.err.println("Error: Duplicate filename detected - " + fileName);
            return; // Skip processing this row
        }

        // Process the remaining columns as the content
        String lineData = generateContent(row);

        // Create the full path for the file
        String filePath = outputFolderPath + "\\" + fileName;

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(new File(filePath)))) {
            writer.write(lineData);
            System.out.println(fileName + " completed successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String generateContent(Row row) {
        StringBuilder templateContent = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(new File(templateFilePath))))) {
            String line;
            while ((line = reader.readLine()) != null) {
                templateContent.append(line).append("\n");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Generate the file content by replacing placeholders with values from the Excel row
        for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
            String placeholder = entry.getKey();
            int columnIndex = entry.getValue();
            Cell cell = row.getCell(columnIndex);
            String cellValue = cell != null ? cellToString(cell) : "";
            templateContent = new StringBuilder(templateContent.toString().replace(placeholder, cellValue));
        }
        return templateContent.toString();
    }

    private static String cellToString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((long) cell.getNumericCellValue());
            default:
                return "";
        }
    }

    private static boolean isValidFilename(String fileName) {
        return filenameSet.add(fileName);
    }

    private static void processTemplate(Row headerRow) {
        // Extract placeholders from the header row and populate headerMap
        int columnCount = headerRow.getLastCellNum();
        for (int i = 0; i < columnCount; i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String placeholder = cell.getStringCellValue();
                headerMap.put(placeholder, i);
            }
        }
    }

    private static String createOutputFolder() throws IOException {
        // Extract the base name of the input file
        String inputFileName = new File(excelFilePath).getName();
        String baseInputFileName = inputFileName.substring(0, inputFileName.lastIndexOf('.'));

        // Create a timestamp for the folder name
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        String timestamp = dateFormat.format(new Date());

        // Create the timestamped output folder
        String outputFolderPath = outputRootFolder + "\\" + baseInputFileName + "_" + timestamp;
        Files.createDirectories(Paths.get(outputFolderPath));
        return outputFolderPath;
    }
}
