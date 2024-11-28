package org.example;

import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.*;

public class Script {
    public static void main(String[] args) throws Exception {
        Scanner scanner = new Scanner(System.in);

        try {
            // Prompt user for input file path
            System.out.println("Provide file path to the Excel file:");
            String filePath = scanner.nextLine();

            // Define the default sheet name
            String sheetName = "work (2)";

            // Read data from the Excel file
            List<List<String>> data = readExcelFile(filePath, sheetName);

            // Process data to keep only unique Defender Atp: Asset Name per inst
            List<List<String>> filteredData = filterUniqueEntries(data);

            // Define output path in the user's Documents folder
            String userDocuments = Paths.get(System.getProperty("user.home"), "Documents").toString();
            String outputFilePath = Paths.get(userDocuments, "FilteredReport.xlsx").toString();

            // Write the filtered data to the output file
            writeToExcel(filteredData, outputFilePath);

            // Notify the user of success
            System.out.println("Filtered report has been created at: " + outputFilePath);
        } catch (Exception e) {
            // Print any errors that occur
            System.err.println("Error processing file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Reads data from the specified Excel sheet.
     *
     * @param filePath  The path to the Excel file.
     * @param sheetName The name of the sheet to read.
     * @return A list of rows, where each row is a list of cell values as strings.
     * @throws Exception If the file cannot be read or the sheet does not exist.
     */
    private static List<List<String>> readExcelFile(String filePath, String sheetName) throws Exception {
        List<List<String>> rows = new ArrayList<>();

        // Open the Excel file
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Get the sheet by name
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet '" + sheetName + "' not found.");
            }

            // Iterate through rows and extract data
            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                row.forEach(cell -> rowData.add(cell.toString().trim()));
                rows.add(rowData);
            }
        }

        return rows;
    }

    /**
     * Filters the data to ensure Defender Atp: Asset Name is unique to one inst.
     *
     * @param data The raw data read from the Excel sheet.
     * @return A filtered list of rows.
     */
    private static List<List<String>> filterUniqueEntries(List<List<String>> data) {
        // Use a set to track asset names that have already been processed
        Set<String> uniqueAssets = new HashSet<>();
        List<List<String>> filteredData = new ArrayList<>();

        // Process the data rows (skip the header row)
        for (int i = 1; i < data.size(); i++) {
            List<String> row = data.get(i);
            String inst = row.get(0); // Column "inst"
            String assetName = row.get(2); // Column "Defender Atp: Asset Name"

            // If the asset name is already in the set, skip this row
            if (!uniqueAssets.add(assetName)) {
                continue;
            }

            // Add the row to the filtered data
            filteredData.add(row);
        }

        return filteredData;
    }

    /**
     * Writes the filtered data to a new Excel file.
     *
     * @param data         The filtered data to write.
     * @param outputPath   The file path to save the output file.
     * @throws Exception If the file cannot be written.
     */
    private static void writeToExcel(List<List<String>> data, String outputPath) throws Exception {
        // Create a new workbook and sheet
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            Sheet sheet = workbook.createSheet("Filtered Data");

            // Write rows to the new sheet
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i);
                List<String> rowData = data.get(i);
                for (int j = 0; j < rowData.size(); j++) {
                    row.createCell(j).setCellValue(rowData.get(j));
                }
            }

            // Write the workbook to the output file
            workbook.write(fos);
        }
    }
}
