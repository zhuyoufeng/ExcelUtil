package com.sargeraswang.util.ExcelUtil;

import org.apache.poi.ss.util.CellRangeAddress;

/**
 * ExcelHeaderCell
 *
 * @author zhuyoufeng
 */
public class ExcelHeaderCell {

    private String text;
    private String property;
    private Integer firstRow;
    private Integer lastRow;
    private Integer firstCol;
    private Integer lastCol;

    public static ExcelHeaderCell createExcelHeaderCell(String text, String property, int firstRow, int lastRow, int firstCol, int lastCol) {
        ExcelHeaderCell excelHeaderCell = new ExcelHeaderCell();
        excelHeaderCell.text = text;
        excelHeaderCell.property = property;
        excelHeaderCell.firstRow = firstRow;
        excelHeaderCell.lastRow = lastRow;
        excelHeaderCell.firstCol = firstCol;
        excelHeaderCell.lastCol = lastCol;
        return excelHeaderCell;
    }

    public static ExcelHeaderCell createExcelHeaderCell(String text, int firstRow, int lastRow, int firstCol, int lastCol) {
        return ExcelHeaderCell.createExcelHeaderCell(text, null, firstRow, lastRow, firstCol, lastCol);
    }

    public static ExcelHeaderCell createExcelHeaderCell(String text, String property, int firstRow, int firstCol) {
        ExcelHeaderCell excelHeaderCell = new ExcelHeaderCell();
        excelHeaderCell.text = text;
        excelHeaderCell.property = property;
        excelHeaderCell.firstRow = firstRow;
        excelHeaderCell.firstCol = firstCol;
        return excelHeaderCell;
    }

    public static ExcelHeaderCell createExcelHeaderCell(String text, int firstRow, int firstCol) {
        ExcelHeaderCell excelHeaderCell = new ExcelHeaderCell();
        excelHeaderCell.text = text;
        excelHeaderCell.firstRow = firstRow;
        excelHeaderCell.firstCol = firstCol;
        return excelHeaderCell;
    }

    public String getText() {
        return text;
    }

    public String getProperty() {
        return property;
    }

    public Integer getFirstRow(int titleRow) {
        return firstRow + titleRow;
    }

    public Integer getFirstCol() {
        return firstCol;
    }

    public Integer getLastRow(int titleRow) {
        return lastRow + titleRow;
    }

    public Integer getLastCol() {
        return lastCol;
    }

    public CellRangeAddress createCellRangeAddress(int titleRow) {
        if (lastRow == null) {
            return null;
        }
        return new CellRangeAddress(firstRow + titleRow, lastRow + titleRow, firstCol, lastCol);
    }
}
