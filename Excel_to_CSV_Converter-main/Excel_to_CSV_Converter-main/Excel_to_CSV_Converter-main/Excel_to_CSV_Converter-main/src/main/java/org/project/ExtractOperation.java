package org.project;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellAddress;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.*;

public class ExtractOperation {

    private static final Logger logger = Logger.getLogger(ExtractOperation.class.getName());

    public static void main(String[] args) {
        configureLogger(); // Configure logger settings

        String excelFilePath = chooseExcelFile();
        if (excelFilePath == null) {
            logError("No Excel file selected.");
            return;
        }

        String outputFolder = chooseOutputFolder();
        if (outputFolder == null) {
            logError("No output folder selected.");
            return;
        }

        String startCell = promptForCell("Enter starting cell (e.g., A1):");
        if (startCell == null || !isValidCellReference(startCell)) {
            logError("Invalid starting cell format.");
            return;
        }

        String endCell = promptForCell("Enter ending cell (e.g., B10):");
        if (endCell != null && !isValidCellReference(endCell)) {
            logError("Invalid ending cell format.");
            return;
        }

        DataExtractor extractor = new DataExtractor(excelFilePath, startCell, endCell, outputFolder);
        extractor.extractAndWriteTransposed();

        System.out.println("Extraction operation completed successfully.");
    }

    private static boolean isValidCellReference(String cellRef) {
        // Basic validation for cell reference format
        return cellRef.matches("[A-Za-z]+\\d+");
    }

    private static String chooseExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Choose Excel File");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx", "xls"));
        int userSelection = fileChooser.showOpenDialog(null);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            return selectedFile.getAbsolutePath();
        }
        return null;
    }

    private static String chooseOutputFolder() {
        JFileChooser folderChooser = new JFileChooser();
        folderChooser.setDialogTitle("Choose Output Folder");
        folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int userSelection = folderChooser.showSaveDialog(null);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File selectedFolder = folderChooser.getSelectedFile();
            return selectedFolder.getAbsolutePath();
        }
        return null;
    }

    private static String promptForCell(String message) {
        return JOptionPane.showInputDialog(message);
    }

    private static void configureLogger() {
        ConsoleHandler consoleHandler = new ConsoleHandler();
        consoleHandler.setLevel(Level.SEVERE);
        logger.addHandler(consoleHandler);

        try {
            FileHandler fileHandler = new FileHandler("extract_operation.log");
            fileHandler.setLevel(Level.ALL);
            SimpleFormatter formatter = new SimpleFormatter();
            fileHandler.setFormatter(formatter);
            logger.addHandler(fileHandler);
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Error configuring logger: " + e.getMessage());
        }
    }

    private static void logError(String message) {
        logger.log(Level.SEVERE, message);
    }

    static class DataExtractor {
        private final String excelFilePath;
        private final String startCellRef;
        private final String endCellRef;
        private final String outputFolder;

        public DataExtractor(String excelFilePath, String startCellRef, String endCellRef, String outputFolder) {
            this.excelFilePath = excelFilePath;
            this.startCellRef = startCellRef;
            this.endCellRef = endCellRef;
            this.outputFolder = outputFolder;
        }

        public void extractAndWriteTransposed() {
            try (FileInputStream fis = new FileInputStream(excelFilePath);
                 Workbook workbook = new XSSFWorkbook(fis)) {

                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    if (sheet == null) {
                        logError("Sheet at index " + i + " not found.");
                        continue;
                    }

                    CellAddress start = new CellAddress(startCellRef);
                    CellAddress end = null;
                    if (endCellRef != null && !endCellRef.trim().isEmpty()) {
                        end = new CellAddress(endCellRef);
                    }

                    List<List<String>> data = new ArrayList<>();

                    // Extract the data
                    int lastRow = end != null ? Math.min(end.getRow(), sheet.getLastRowNum()) : sheet.getLastRowNum();
                    for (int row = start.getRow(); row <= lastRow; row++) {
                        Row currentRow = sheet.getRow(row);
                        if (currentRow != null) {
                            boolean skipRow = false;
                            List<String> rowData = new ArrayList<>();
                            for (int col = start.getColumn(); col <= currentRow.getLastCellNum(); col++) {
                                Cell cell = currentRow.getCell(col);
                                // Check if the cell contains "Comment" string
                                if (cell != null && cell.getCellType() == CellType.STRING) {
                                    String cellValue = cell.getStringCellValue().trim();
                                    if ("Comment".equalsIgnoreCase(cellValue)) {
                                        skipRow = true;
                                        break; // Skip this entire row
                                    }
                                }
                                rowData.add(getCellValueAsString(cell));
                            }
                            if (!skipRow && !rowData.isEmpty()) {
                                data.add(rowData);
                            }
                        }
                    }

                    // Transpose the data
                    List<List<String>> transposedData = transposeData(data);

                    // Write the transposed data to CSV
                    String outputFileName = sheet.getSheetName() + ".csv";
                    File csvFile = new File(outputFolder, outputFileName);

                    try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
                        for (List<String> rowData : transposedData) {
                            String rowString = String.join(",", rowData);
                            bw.write(rowString);
                            bw.newLine();
                        }
                    }

                    logger.info("Extracting and transposing data from " + startCellRef + " to " + endCellRef + " in sheet " + sheet.getSheetName() + " completed successfully.");
                }

            } catch (IOException e) {
                logError("Error processing Excel file: " + e.getMessage());
            }
        }

        private List<List<String>> transposeData(List<List<String>> data) {
            if (data.isEmpty() || data.get(0).isEmpty()) {
                return new ArrayList<>();
            }
            int rowCount = data.size();
            int colCount = data.get(0).size();
            List<List<String>> transposedData = new ArrayList<>(colCount);

            for (int col = 0; col < colCount; col++) {
                List<String> transposedRow = new ArrayList<>(rowCount);
                for (List<String> rowData : data) {
                    transposedRow.add(rowData.size() > col ? rowData.get(col) : "");
                }
                transposedData.add(transposedRow);
            }
            return transposedData;
        }

        private String getCellValueAsString(Cell cell) {
            if (cell == null) {
                return "";
            }
            switch (cell.getCellType()) {
                case STRING:
                    String cellValue = cell.getStringCellValue().replace("\n", " ").replace("\r", "").trim();
                    // Check if the cell value contains a comma
                    if (cellValue.contains(",")) {
                        // Enclose the cell value in double quotes to preserve commas in CSV
                        cellValue = "\"" + cellValue + "\"";
                    }
                    return cellValue;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    return cell.getCellFormula();
                case BLANK:
                default:
                    return "";
            }
        }
    }
}
