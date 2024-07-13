package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.logging.*;

public class RotateOperation {

    private static final Logger logger = Logger.getLogger(RotateOperation.class.getName());

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

        int rotationDegree = promptForRotationDegree();
        if (rotationDegree == -1) {
            logError("Invalid rotation degree entered.");
            return;
        }

        rotateAndWrite(excelFilePath, rotationDegree, outputFolder);

        System.out.println("Rotating operation completed successfully.");
    }

    private static void rotateAndWrite(String excelFilePath, int degree, String outputFolder) {
        try (FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();

                int[][] matrix = createMatrix(sheet);
                int[][] rotatedMatrix = rotateMatrix(matrix, degree);

                // Adjusting file name to include rotation degree and keep original sheet name
                String outputFileName = "A_" + sheetName + "_rotated_" + degree + ".csv";
                File csvFile = new File(outputFolder, outputFileName);

                try (BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(csvFile), StandardCharsets.UTF_8))) {
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
                                try {
                                    matrix[row][col] = Integer.parseInt(cell.getStringCellValue());
                                } catch (NumberFormatException e) {
                                    matrix[row][col] = 0;
                                }
                                break;
                            case NUMERIC:
                                matrix[row][col] = (int) cell.getNumericCellValue();
                                break;
                            case BOOLEAN:
                                matrix[row][col] = cell.getBooleanCellValue() ? 1 : 0;
                                break;
                            case FORMULA:
                                matrix[row][col] = (int) cell.getNumericCellValue();
                                break;
                            case BLANK:
                            default:
                                matrix[row][col] = 0;
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
                rotatedMatrix = matrix;
                break;
        }

        return rotatedMatrix;
    }

    private static void logError(String message) {
        logger.log(Level.SEVERE, message);
    }

    private static void configureLogger() {
        ConsoleHandler consoleHandler = new ConsoleHandler();
        consoleHandler.setLevel(Level.SEVERE);
        logger.addHandler(consoleHandler);

        try {
            FileHandler fileHandler = new FileHandler("rotate_operation.log");
            fileHandler.setLevel(Level.ALL);
            SimpleFormatter formatter = new SimpleFormatter();
            fileHandler.setFormatter(formatter);
            logger.addHandler(fileHandler);
        } catch (IOException e) {
            logger.log(Level.SEVERE, "Error configuring logger: " + e.getMessage());
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

    private static int promptForRotationDegree() {
        String degreeString = JOptionPane.showInputDialog("Enter rotation degree (90, 180, 270):");
        try {
            int degree = Integer.parseInt(degreeString);
            if (degree == 90 || degree == 180 || degree == 270) {
                return degree;
            } else {
                logError("Invalid rotation degree. Must be 90, 180, or 270.");
                return -1;
            }
        } catch (NumberFormatException e) {
            logError("Invalid rotation degree. Must be an integer.");
            return -1;
        }
    }
}
