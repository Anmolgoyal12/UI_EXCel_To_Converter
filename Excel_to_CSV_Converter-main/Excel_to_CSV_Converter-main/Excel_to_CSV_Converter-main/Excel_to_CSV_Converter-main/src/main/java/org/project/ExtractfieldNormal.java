package org.project;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellAddress;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.logging.*;
import java.awt.event.KeyEvent;
import java.awt.event.KeyAdapter;

public class ExtractfieldNormal {
    private static final Logger logger = Logger.getLogger(ExtractfieldNormal.class.getName());
    private static boolean isEscPressed = false;

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

        String startCell = promptForCell();
        if (startCell == null) {
            logError("Invalid starting cell.");
            return;
        }

        String endCell = promptForEndCell();
        if (isEscPressed) {
            endCell = null;
        }

        String sheetName = chooseSheetName(excelFilePath);
        if (sheetName == null) {
            logError("No sheet selected.");
            return;
        }

        extractAndWrite(excelFilePath, startCell, endCell, outputFolder, sheetName);

        System.out.println("Extraction operation completed successfully.");
    }

    private static void extractAndWrite(String excelFilePath, String startCellRef, String endCellRef, String outputFolder, String sheetName) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                logError("Sheet " + sheetName + " not found.");
                return;
            }

            CellAddress start = new CellAddress(startCellRef);
            CellAddress end = (endCellRef != null && !endCellRef.trim().isEmpty()) ? new CellAddress(endCellRef) : null;

            String outputFileName = sheetName + ".csv";
            File csvFile = new File(outputFolder, outputFileName);

            try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
                int endRow = (end != null) ? end.getRow() : sheet.getLastRowNum();
                int endCol = (end != null) ? end.getColumn() : -1; // To handle full row extraction

                Integer commentColIndex = null;

                for (int row = start.getRow(); row <= endRow; row++) {
                    Row currentRow = sheet.getRow(row);
                    if (currentRow != null) {
                        if (row == start.getRow()) { // Check headers only for the first row
                            for (int col = start.getColumn(); col < currentRow.getLastCellNum(); col++) {
                                Cell cell = currentRow.getCell(col);
                                if (cell != null && cell.getCellType() == CellType.STRING) {
                                    String cellValue = cell.getStringCellValue().trim();
                                    if ("Comment".equalsIgnoreCase(cellValue)) {
                                        commentColIndex = col; // Set comment column index
                                        break;
                                    }
                                }
                            }
                        }

                        StringBuilder rowString = new StringBuilder();
                        for (int col = start.getColumn(); col <= (endCol != -1 ? endCol : currentRow.getLastCellNum() - 1); col++) {
                            if (commentColIndex != null && col >= commentColIndex) break; // Skip columns starting from the comment column

                            Cell cell = currentRow.getCell(col);
                            String cellValue = getCellValueAsString(cell);
                            if (rowString.length() > 0) {
                                rowString.append(",");
                            }
                            rowString.append(escapeCsvValue(cellValue));
                        }

                        // Write row to CSV if it has content
                        if (rowString.length() > 0) {
                            bw.write(rowString.toString());
                            bw.newLine();
                        }
                    }
                }
            }

            logger.info("Extracting data from " + startCellRef + " to " + (endCellRef != null ? endCellRef : "end") + " completed successfully.");

        } catch (IOException e) {
            logError("Error processing Excel file: " + e.getMessage());
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().replace("\n", " ");
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

    private static String promptForCell() {
        return JOptionPane.showInputDialog("Enter starting cell (e.g., A1):");
    }

    private static String promptForEndCell() {
        JTextField textField = new JTextField();
        textField.addKeyListener(new KeyAdapter() {
            public void keyPressed(KeyEvent e) {
                if (e.getKeyCode() == KeyEvent.VK_ESCAPE) {
                    isEscPressed = true;
                    ((JOptionPane) textField.getParent().getParent()).setValue(JOptionPane.CLOSED_OPTION);
                }
            }
        });
        int option = JOptionPane.showConfirmDialog(null, textField, "Enter ending cell (e.g., B10) or press ESC to skip:", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (option == JOptionPane.OK_OPTION) {
            return textField.getText();
        } else {
            return null;
        }
    }

    private static String chooseSheetName(String excelFilePath) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            String[] sheetNames = new String[workbook.getNumberOfSheets()];
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames[i] = workbook.getSheetName(i);
            }
            return (String) JOptionPane.showInputDialog(null, "Choose sheet",
                    "Sheet Selection", JOptionPane.QUESTION_MESSAGE, null, sheetNames, sheetNames[0]);
        } catch (IOException e) {
            logError("Error reading Excel file: " + e.getMessage());
            return null;
        }
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

    private static String escapeCsvValue(String value) {
        if (value.contains(",") || value.contains("\"") || value.contains("\n")) {
            value = value.replace("\"", "\"\"");
            return "\"" + value + "\"";
        } else {
            return value;
        }
    }
}
