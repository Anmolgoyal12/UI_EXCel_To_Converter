package org.csdconverter;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This class converts data from Excel sheets to CSV files based on
 * configuration.
 */
public class MainCSD {

    private static final String BASE_OUTPUT_DIR = "D:/Excel_to_CSV_Converter-main/BASE_OUTPUT_DIRECTORY";
    private static final Logger logger = Logger.getLogger(MainCSD.class.getName());

    /**
     * Main method to initiate the Excel to CSV conversion process.
     *
     * @param args Command-line arguments (not used in this application)
     */
    public static void main(String[] args) {
        MainUI.launch(MainUI.class, args);
    }

    public static void convert(String configFilePath, String excelFilePath) {
        List<SheetConfig> sheetConfigs = loadSheetConfigs(configFilePath);

        for (SheetConfig config : sheetConfigs) {
            String sheetName = config.getSheetName();
            String csvFilePath = Paths.get(BASE_OUTPUT_DIR, config.getOutputDirectory(), config.getCsvName()).toString();
            try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath))) {
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet != null) {
                    List<List<String>> extractedData = extractDataFromSheet(sheet, config);

                    if (config.isTranspose() && !config.getExcludeFromTranspose().contains(sheetName)) {
                        extractedData = transposeData(extractedData);
                    }

                    applyAdvanceConditionToHeaders(extractedData);
                    cleanUpData(extractedData);
                    writeCSV(csvFilePath, extractedData);
                } else {
                    throw new Exception("Sheet not found: " + sheetName);
                }
            } catch (Exception e) {
                logger.severe("Error processing sheet: " + sheetName + ". " + e.getMessage());
            }
        }

        logger.info("Conversion completed successfully.");
    }

    public static List<SheetConfig> loadSheetConfigs(String configFilePath) {
        List<SheetConfig> sheetConfigs = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(configFilePath))) {
            Sheet configSheet = workbook.getSheetAt(0);
            for (Row row : configSheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                SheetConfig config = new SheetConfig(
                        getCellValue(row.getCell(1)),
                        getCellValue(row.getCell(2)),
                        getTextBooleanCellValue(row.getCell(3)),
                        getTextBooleanCellValue(row.getCell(4)),
                        getCellValue(row.getCell(5)),
                        getStringListCellValue(row.getCell(6)),
                        getCellValue(row.getCell(7))
                );
                sheetConfigs.add(config);
            }
        } catch (IOException e) {
            logger.severe("Error loading sheet configurations: " + e.getMessage());
        }
        return sheetConfigs;
    }

    public static String standardizeHeader(String input) {
        if (input == null || input.isEmpty()) {
            return input;
        }

        input = input.replace("*", "");
        while (input.endsWith("_")) {
            input = input.substring(0, input.length() - 1);
        }

        String lowerCaseInput = input.toLowerCase();
        if ("username".equalsIgnoreCase(lowerCaseInput)) {
            return "user_name";
        }

        return lowerCaseInput.replaceAll("\\s+", "_");
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static boolean getTextBooleanCellValue(Cell cell) {
        if (cell == null) {
            return false;
        }

        if (cell.getCellType() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().trim();
            return "true".equalsIgnoreCase(cellValue);
        }

        return false;
    }

    private static List<String> getStringListCellValue(Cell cell) {
        List<String> stringList = new ArrayList<>();
        if (cell != null && cell.getCellType() == CellType.STRING) {
            String[] values = cell.getStringCellValue().split(",");
            for (String value : values) {
                stringList.add(value.trim());
            }
        }
        return stringList;
    }

    private static List<List<String>> extractDataFromSheet(Sheet sheet, SheetConfig config) {
        List<List<String>> data = new ArrayList<>();

        boolean shouldTranspose = config.isTranspose();
        logger.info("Sheet: " + sheet.getSheetName() + " - Should Transpose: " + shouldTranspose);
        int startColumn = 1;

        Row headerRow = sheet.getRow(0);
        int commentColumnIndex = -1;

        if (headerRow != null) {
            for (Cell cell : headerRow) {
                if ("Comment".equals(cell.getStringCellValue()) || "Comments".equals(cell.getStringCellValue())) {
                    commentColumnIndex = cell.getColumnIndex();
                    break;
                }
            }
        }

        String range = config.getRange();
        int startRow = shouldTranspose ? 2 : 0;

        List<Integer> rowIndices = new ArrayList<>();

        if (range != null && !range.isEmpty() && !"NA".equalsIgnoreCase(range)) {
            String[] parts = range.split(",");
            for (String part : parts) {
                if (part.contains("-")) {
                    String[] bounds = part.split("-");
                    try {
                        int start = Integer.parseInt(bounds[0].trim());
                        int end = Integer.parseInt(bounds[1].trim());
                        for (int i = start; i <= end; i++) {
                            rowIndices.add(i - 1);
                        }
                    } catch (NumberFormatException e) {
                        logger.severe("Invalid range format: " + range);
                    }
                } else if (part.matches("\\d+")) {
                    int row = Integer.parseInt(part.trim()) - 1;
                    rowIndices.add(row);
                } else if (part.matches("[A-Z]+\\d+")) {
                    CellReference cellReference = new CellReference(part.trim());
                    int row = cellReference.getRow();
                    rowIndices.add(row);
                } else {
                    logger.severe("Invalid range format: " + range);
                }
            }
        }

        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            if (!rowIndices.isEmpty() && !rowIndices.contains(i)) {
                continue;
            }

            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }

            if (config.isCommentRead() != null && config.isCommentRead()) {
                Cell firstCell = row.getCell(0);
                String cellValue = getCellValue(firstCell);
                if (cellValue != null && cellValue.startsWith("#")) {
                    continue;
                }
            }
            List<String> rowData = new ArrayList<>();
            for (int j = startColumn; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (!config.isCommentRead() && j == commentColumnIndex) {
                    continue;
                }
                rowData.add(getCellValue(cell));
            }

            data.add(rowData);
        }

        return data;
    }

    private static List<List<String>> transposeData(List<List<String>> data) {
        List<List<String>> transposedData = new ArrayList<>();
        if (data.isEmpty() || data.get(0).isEmpty()) {
            return transposedData;
        }

        int colCount = data.get(0).size();

        for (int col = 0; col < colCount; col++) {
            List<String> transposedRow = new ArrayList<>();
            for (List<String> currentRow : data) {
                if (col < currentRow.size()) {
                    transposedRow.add(currentRow.get(col));
                } else {
                    transposedRow.add("");
                }
            }
            transposedData.add(transposedRow);
        }

        return transposedData;
    }

    private static void applyAdvanceConditionToHeaders(List<List<String>> data) {
        if (data == null || data.isEmpty()) {
            return;
        }

        List<String> headers = data.get(0);
        headers.replaceAll(MainCSD::standardizeHeader);
    }

    private static String escapeCsvData(String data) {
        if (data.contains(",") || data.contains("\n") || data.contains("\"")) {
            data = data.replace("\"", "\"\"");
            data = "\"" + data + "\"";
        }
        return data.toLowerCase();
    }

    private static void cleanUpData(List<List<String>> data) {
        for (List<String> row : data) {
            row.replaceAll(s -> s.replace("*", ""));
        }
    }

    private static void writeCSV(String csvFilePath, List<List<String>> data) {
        File outputFile = new File(csvFilePath);
        if (!outputFile.getParentFile().exists() && !outputFile.getParentFile().mkdirs()) {
            logger.severe("Failed to create output directories for: " + csvFilePath);
            return;
        }
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile))) {
            for (List<String> row : data) {
                String line = row.stream().map(MainCSD::escapeCsvData).collect(Collectors.joining(","));
                writer.write(line);
                writer.newLine();
            }
        } catch (IOException e) {
            logger.severe("Error writing CSV file: " + csvFilePath + ". " + e.getMessage());
        }
    }
}
