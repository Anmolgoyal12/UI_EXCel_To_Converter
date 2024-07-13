package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellAddress;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.logging.*;

public class normalcsv {

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
        if (startCell == null) {
            logError("Invalid starting cell.");
            return;
        }

        String endCell = promptForCell("Enter ending cell (e.g., B10):");
        if (endCell == null) {
            logError("Invalid ending cell.");
            return;
        }

        extractAndWrite(excelFilePath, startCell, endCell, outputFolder);

        System.out.println("Extraction operation completed successfully.");
    }

    private static void extractAndWrite(String excelFilePath, String startCellRef, String endCellRef, String outputFolder) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                CellAddress start = new CellAddress(startCellRef);
                CellAddress end = new CellAddress(endCellRef);

                String outputFileName = sheetName + ".csv";
                File csvFile = new File(outputFolder, outputFileName);

                try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
                    for (int row = start.getRow(); row <= end.getRow(); row++) {
                        StringBuilder rowString = new StringBuilder();
                        Row currentRow = sheet.getRow(row);
                        if (currentRow != null) {
                            for (int col = start.getColumn(); col <= end.getColumn(); col++) {
                                Cell cell = currentRow.getCell(col);
                                // Ignore columns with the "Comment" field
                                if (cell != null && cell.getCellType() == CellType.STRING && "Comment".equalsIgnoreCase(cell.getStringCellValue())) {
                                    continue;
                                }
                                if (rowString.length() > 0) {
                                    rowString.append(",");
                                }
                                rowString.append(getCellValueAsString(cell));
                            }
                        }
                        if (rowString.length() > 0) {
                            bw.write(rowString.toString());
                            bw.newLine();
                        }
                    }
                }
            }

            logger.info("Extracting data from " + startCellRef + " to " + endCellRef + " completed successfully.");

        } catch (IOException e) {
            logError("Error processing Excel file: " + e.getMessage());
        }
    }

    private static void logError(String message) {
        logger.log(Level.SEVERE, message);
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

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
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
}
