package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.Scanner;
import java.util.logging.*;

public class MannualPromptExtraction {

    // Logger for logging errors and informational messages
    private static final Logger logger = Logger.getLogger(MannualPromptExtraction.class.getName());

    public static void main(String[] args) {
        configureLogger(); // Configure logger settings
        String excelFilePath = chooseExcelFile();
        if (excelFilePath == null) {
            logError("No Excel file selected or access denied.");
            return;
        }

        String operation = getUserInput("Enter the operation (transpose/rotate/extract): ").toLowerCase();
        if (!isValidOperation(operation)) {
            logError("Invalid operation. Supported operations are transpose, rotate, and extract.");
            return;
        }

        String outputFolder = chooseOutputFolder();
        if (outputFolder == null) {
            logError("No output folder selected.");
            return;
        }

        switch (operation) {
            case "rotate":
                int degree = getRotationDegree();
                rotateAndWrite(excelFilePath, degree, outputFolder);
                break;
            case "extract":
                String startCell = getUserInput("Enter the starting cell (e.g., A1): ");
                String endCell = getUserInput("Enter the ending cell (e.g., C3): ");
                extractCellsAndWrite(excelFilePath, startCell, endCell, outputFolder);
                break;
            case "transpose":
                transposeAndWrite(excelFilePath, outputFolder);
                break;
            default:
                logError("Invalid operation. Supported operations are transpose, rotate, and extract.");
                break;
        }

        System.out.println("Excel file operation completed successfully.");
    }

    private static void transposeAndWrite(String excelFilePath, String outputFolder) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                File csvFile = new File(outputFolder, sheetName + "_transposed.csv");

                try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
                    // Transpose logic
                    for (int col = 0; col < sheet.getRow(0).getLastCellNum(); col++) {
                        StringBuilder rowString = new StringBuilder();
                        for (int row = 0; row <= sheet.getLastRowNum(); row++) {
                            Row currentRow = sheet.getRow(row);
                            if (currentRow != null && currentRow.getCell(col) != null) {
                                if (rowString.length() > 0) {
                                    rowString.append(",");
                                }
                                rowString.append(getCellValueAsString(currentRow.getCell(col)));
                            }
                        }
                        bw.write(rowString.toString());
                        bw.newLine();
                    }
                }
            }

            logger.info("Transposing Excel file completed successfully.");

        } catch (IOException e) {
            logError("Error processing Excel file: " + e.getMessage());
        }
    }

    private static void rotateAndWrite(String excelFilePath, int degree, String outputFolder) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                int[][] matrix = createMatrix(sheet);
                int[][] rotatedMatrix = rotateMatrix(matrix, degree);

                File csvFile = new File(outputFolder, sheetName + "_rotated_" + degree + ".csv");

                try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
                    // Write rotated matrix to CSV
                    for (int[] row : rotatedMatrix) {
                        StringBuilder rowString = new StringBuilder();
                        for (int cell : row) {
                            if (rowString.length() > 0) {
                                rowString.append(",");
                            }
                            rowString.append(cell);
                        }
                        bw.write(rowString.toString());
                        bw.newLine();
                    }
                }
            }

            logger.info("Rotating Excel file by " + degree + " degrees completed successfully.");

        } catch (IOException e) {
            logError("Error processing Excel file: " + e.getMessage());
        }
    }

    private static void extractCellsAndWrite(String excelFilePath, String startCell, String endCell, String outputFolder) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                CellReference startRef = new CellReference(startCell);
                CellReference endRef = new CellReference(endCell);

                int startRow = startRef.getRow();
                int endRow = endRef.getRow();
                int startCol = startRef.getCol();
                int endCol = endRef.getCol();

                File csvFile = new File(outputFolder, sheetName + "_extracted_" + startCell + "_" + endCell + ".csv");

                try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
                    // Write cells in the specified range to CSV
                    for (int row = startRow; row <= endRow; row++) {
                        Row currentRow = sheet.getRow(row);
                        if (currentRow != null) {
                            StringBuilder rowString = new StringBuilder();
                            for (int col = startCol; col <= endCol; col++) {
                                Cell cell = currentRow.getCell(col);
                                if (cell != null) {
                                    if (rowString.length() > 0) {
                                        rowString.append(",");
                                    }
                                    rowString.append(getCellValueAsString(cell));
                                }
                            }
                            bw.write(rowString.toString());
                            bw.newLine();
                        }
                    }
                }
            }

            logger.info("Extracting cells from Excel file completed successfully.");

        } catch (IOException e) {
            logError("Error processing Excel file: " + e.getMessage());
        }
    }

    private static int[][] createMatrix(Sheet sheet) {
        int numRows = sheet.getLastRowNum() + 1;
        int numCols = sheet.getRow(0).getLastCellNum();
        int[][] matrix = new int[numRows][numCols];

        for (int row = 0; row < numRows; row++) {
            Row currentRow = sheet.getRow(row);
            if (currentRow != null) {
                for (int col = 0; col < numCols; col++) {
                    Cell cell = currentRow.getCell(col);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                // Handle string values if needed
                                String stringValue = cell.getStringCellValue();
                                // Convert string value to numeric if applicable
                                try {
                                    matrix[row][col] = Integer.parseInt(stringValue);
                                } catch (NumberFormatException e) {
                                    // Handle if string cannot be parsed to integer
                                    matrix[row][col] = 0; // or another default value
                                }
                                break;
                            case NUMERIC:
                                // Check if the cell is formatted as numeric
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    matrix[row][col] = (int) cell.getDateCellValue().getTime();
                                } else {
                                    matrix[row][col] = (int) cell.getNumericCellValue();
                                }
                                break;
                            case BOOLEAN:
                                // Handle boolean values if needed
                                matrix[row][col] = cell.getBooleanCellValue() ? 1 : 0;
                                break;
                            case FORMULA:
                                // Handle formulas if needed
                                matrix[row][col] = (int) cell.getNumericCellValue(); // or another appropriate handling
                                break;
                            case BLANK:
                            default:
                                // Handle blank cells or other types
                                matrix[row][col] = 0; // or another default value
                                break;
                        }
                    }
                }
            }
        }
        return matrix;
    }

    private static int[][] rotateMatrix(int[][] matrix, int degree) {
        int numRows = matrix.length;
        int numCols = matrix[0].length;
        int[][] rotatedMatrix;

        switch (degree % 360) {
            case 90:
                rotatedMatrix = new int[numCols][numRows];
                for (int i = 0; i < numRows; i++) {
                    for (int j = 0; j < numCols; j++) {
                        rotatedMatrix[j][numRows - 1 - i] = matrix[i][j];
                    }
                }
                break;
            case 180:
                rotatedMatrix = new int[numRows][numCols];
                for (int i = 0; i < numRows; i++) {
                    for (int j = 0; j < numCols; j++) {
                        rotatedMatrix[numRows - 1 - i][numCols - 1 - j] = matrix[i][j];
                    }
                }
                break;
            case 270:
                rotatedMatrix = new int[numCols][numRows];
                for (int i = 0; i < numRows; i++) {
                    for (int j = 0; j < numCols; j++) {
                        rotatedMatrix[numCols - 1 - j][i] = matrix[i][j];
                    }
                }
                break;
            default:
                rotatedMatrix = matrix; // No rotation
                break;
        }

        return rotatedMatrix;
    }

    private static int getRotationDegree() {
        Scanner scanner = new Scanner(System.in);
        System.out.print("Enter the degree to rotate (-∞ to +∞): ");
        try {
            return Integer.parseInt(scanner.nextLine().trim());
        } catch (NumberFormatException e) {
            logError("Invalid input. Please enter a valid degree.");
            return -1;
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

    private static String getUserInput(String message) {
        Scanner scanner = new Scanner(System.in);
        System.out.print(message);
        return scanner.nextLine().trim();
    }

    private static boolean isValidOperation(String operation) {
        return operation.equals("transpose") || operation.equals("rotate") || operation.equals("extract");
    }

    private static void logError(String message) {
        logger.log(Level.SEVERE, message);
    }

    private static void configureLogger() {
        // Configure logger to output to console and file
        ConsoleHandler consoleHandler = new ConsoleHandler();
        consoleHandler.setLevel(Level.SEVERE);
        logger.addHandler(consoleHandler);

        try {
            FileHandler fileHandler = new FileHandler("converter.log");
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
}
