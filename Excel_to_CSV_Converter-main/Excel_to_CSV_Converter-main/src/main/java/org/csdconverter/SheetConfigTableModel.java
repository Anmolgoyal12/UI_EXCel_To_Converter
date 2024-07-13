package org.csdconverter;

public class SheetConfigTableModel {

    private final String sheetName;
    private final String csvName;
    private final String isTranspose;

    public SheetConfigTableModel(SheetConfig config) {
        this.sheetName = config.getSheetName();
        this.csvName = config.getCsvName();
        this.isTranspose = config.isTranspose() ? "Yes" : "No";
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getCsvName() {
        return csvName;
    }

    public String isTranspose() {
        return isTranspose;
    }
}
