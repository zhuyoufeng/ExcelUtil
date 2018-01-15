package com.sargeraswang.util.ExcelUtil;

import org.apache.commons.beanutils.BeanComparator;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.ComparatorUtils;
import org.apache.commons.collections.comparators.ComparableComparator;
import org.apache.commons.collections.comparators.ComparatorChain;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * ExcelExportUtil
 *
 * @author zhuyoufeng
 */
public class ExcelExportUtil extends ExcelUtil {

    private static Logger LG = LoggerFactory.getLogger(ExcelExportUtil.class);

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br> 用于单个sheet
     *
     * @param headers 表格属性列名数组
     * @param dataset 需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的 javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    public static <T> void exportExcel(LinkedHashMap<String, String> headers, Collection<T> dataset, OutputStream out) {
        exportExcel(null, null, headers, dataset, out, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br> 用于单个sheet
     *
     * @param headers 表格属性列名数组
     * @param dataset 需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的 javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     * @param pattern 如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    public static <T> void exportExcel(String sheetName, String title, LinkedHashMap<String, String> headers, Collection<T> dataset, OutputStream out, String pattern) {
        try {
            // 声明一个工作薄
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 生成一个表格
            HSSFSheet sheet = (sheetName != null && sheetName.length() > 0) ? workbook.createSheet(sheetName) : workbook.createSheet();
            write2Sheet(workbook, sheet, title, headers, dataset, pattern);
            workbook.write(out);
        } catch (IOException e) {
            LG.error(e.toString(), e);
        }
    }

    public static <T> void exportExcel(List<ExcelHeaderCell> headers, Integer headerSize, Collection<T> dataset, OutputStream out, String pattern) {
        exportExcel(null, null, headers, headerSize, dataset, out, pattern);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br> 用于单个sheet
     *
     * @param headers 表格属性列名数组
     * @param dataset 需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的 javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     * @param pattern 如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    public static <T> void exportExcel(String sheetName, String title, List<ExcelHeaderCell> headers, Integer headerSize, Collection<T> dataset, OutputStream out, String pattern) {
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = (sheetName != null && sheetName.length() > 0) ? workbook.createSheet(sheetName) : workbook.createSheet();
        write2Sheet(workbook, sheet, title, headers, headerSize, dataset, pattern);
        try {
            workbook.write(out);
        } catch (IOException e) {
            LG.error(e.toString(), e);
        }
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br> 用于多个sheet
     *
     * @param sheets {@link ExcelSheet}的集合
     * @param out    与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    public static <T> void exportExcel(List<ExcelSheet<T>> sheets, OutputStream out) {
        exportExcel(sheets, out, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上<br> 用于多个sheet
     *
     * @param sheets  {@link ExcelSheet}的集合
     * @param out     与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     * @param pattern 如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    public static <T> void exportExcel(List<ExcelSheet<T>> sheets, OutputStream out, String pattern) {
        if (CollectionUtils.isEmpty(sheets)) {
            return;
        }
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        for (ExcelSheet<T> sheet : sheets) {
            // 生成一个表格
            HSSFSheet hssfSheet = workbook.createSheet(sheet.getSheetName());
            if (sheet.getMapHeaders() != null) {
                write2Sheet(workbook, hssfSheet, null, sheet.getMapHeaders(), sheet.getDataset(), pattern);
            } else if (sheet.getDefinedHeaders() != null) {
                write2Sheet(workbook, hssfSheet, null, sheet.getDefinedHeaders(), sheet.getHeaderSize(), sheet.getDataset(), pattern);
            }
        }
        try {
            workbook.write(out);
        } catch (IOException e) {
            LG.error(e.toString(), e);
        }
    }

    /**
     * 每个sheet的写入
     *
     * @param sheet   页签
     * @param headers 表头
     * @param dataset 数据集合
     * @param pattern 日期格式
     */
    private static <T> void write2Sheet(HSSFWorkbook workbook, HSSFSheet sheet, String title, LinkedHashMap<String, String> headers, Collection<T> dataset, String pattern) {
        //时间格式默认"yyyy-MM-dd"
        if (StringUtils.isEmpty(pattern)) {
            pattern = "yyyy-MM-dd";
        }
        HSSFCellStyle contentStyle = ExcelStyleUtil.createContentStyle(workbook);
        // 产生表格标题行
        int titleRow = setupTitleRows(workbook, sheet, title, headers.size());
        //标题列数
        int headerRow = setupHeaderRows(workbook, sheet, headers, titleRow);
        // 遍历集合数据，产生数据行
        setupContentRows(workbook, sheet, headers, dataset, pattern, contentStyle, titleRow + headerRow);
        // 设定自动宽度
        autoSizeColumns(sheet, headers.size());
    }

    /**
     * 每个sheet的写入
     *
     * @param sheet   页签
     * @param headers 表头
     * @param dataset 数据集合
     * @param pattern 日期格式
     */
    private static <T> void write2Sheet(HSSFWorkbook workbook, HSSFSheet sheet, String title, List<ExcelHeaderCell> headers, Integer headerSize, Collection<T> dataset, String pattern) {
        //时间格式默认"yyyy-MM-dd"
        if (StringUtils.isEmpty(pattern)) {
            pattern = "yyyy-MM-dd";
        }
        // 产生表格标题行
        int titleRow = setupTitleRows(workbook, sheet, title, headerSize);
        // 标题行
        int headerRow = setupHeaderRows(workbook, sheet, headers, headerSize, titleRow);
        // 遍历集合数据，产生数据行
        setupContentRows(workbook, sheet, headers, dataset, pattern, (headerRow + titleRow));
        // 设定自动宽度
        autoSizeColumns(sheet, headerSize);
    }

    private static void autoSizeColumns(HSSFSheet sheet, Integer headerSize) {
        for (int i = 0; i < headerSize; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private static int setupHeaderRows(HSSFWorkbook workbook, HSSFSheet sheet, List<ExcelHeaderCell> headers, Integer headerSize, int titleRow) {
        int headerRow = 1;
        HSSFCellStyle headerStyle = ExcelStyleUtil.createHeaderStyle(workbook);
        Map<Integer, HSSFRow> storageMap = new HashMap<>();
        for (ExcelHeaderCell excelHeaderCell : headers) {
            CellRangeAddress cellRangeAddress = excelHeaderCell.createCellRangeAddress(titleRow);
            if (cellRangeAddress != null) {
                sheet.addMergedRegion(cellRangeAddress);
                int temp = (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()) + 1;
                if (temp > headerRow) {
                    headerRow = temp;
                }
            }
            HSSFRow tempRow = storageMap.get(excelHeaderCell.getFirstRow(titleRow));
            if (tempRow == null) {
                tempRow = sheet.createRow(excelHeaderCell.getFirstRow(titleRow));
                storageMap.put(excelHeaderCell.getFirstRow(titleRow), tempRow);
            }
            HSSFCell cell = tempRow.createCell(excelHeaderCell.getFirstCol());
            cell.setCellValue(new HSSFRichTextString(excelHeaderCell.getText()));
        }
        for (int i = titleRow; i < headerRow + titleRow; i++) {
            HSSFRow tempRow = sheet.getRow(i);
            for (int j = 0; j < headerSize; j++) {
                HSSFCell cell = tempRow.getCell(j);
                if (cell == null) {
                    cell = tempRow.createCell(j);
                }
                cell.setCellStyle(headerStyle);
            }
        }
        return headerRow;
    }

    private static int setupHeaderRows(HSSFWorkbook workbook, HSSFSheet sheet, LinkedHashMap<String, String> headers, int titleRow) {
        int c = 0;
        HSSFCellStyle headerStyle = ExcelStyleUtil.createHeaderStyle(workbook);
        HSSFRow row = sheet.createRow(titleRow);
        for (Map.Entry<String, String> entry : headers.entrySet()) {
            HSSFCell cell = row.createCell(c);
            cell.setCellStyle(headerStyle);
            HSSFRichTextString text = new HSSFRichTextString(entry.getValue());
            cell.setCellValue(text);
            c++;
        }
        return 1;
    }

    private static int setupTitleRows(HSSFWorkbook workbook, HSSFSheet sheet, String title, Integer headerSize) {
        HSSFCellStyle titleStyle = ExcelStyleUtil.createTitleStyle(workbook);
        int titleRow = 0;
        if (StringUtils.isNotEmpty(title)) {
            titleRow = 1;
            HSSFRow row = sheet.createRow(0);
            if (headerSize > 1) {
                CellRangeAddress rangeAddress = new CellRangeAddress(0, 0, 0, headerSize - 1);
                sheet.addMergedRegion(rangeAddress);
            }
            HSSFCell cell = row.createCell(0);
            cell.setCellStyle(titleStyle);
            HSSFRichTextString text = new HSSFRichTextString(title);
            cell.setCellValue(text);

            for (int i = 0; i < titleRow; i++) {
                HSSFRow tempRow = sheet.getRow(i);
                for (int j = 0; j < headerSize; j++) {
                    HSSFCell tempCell = tempRow.getCell(j);
                    if (tempCell == null) {
                        tempCell = tempRow.createCell(j);
                    }
                    tempCell.setCellStyle(titleStyle);
                }
            }
        }
        return titleRow;
    }

    private static <T> void setupContentRows(HSSFWorkbook workbook, HSSFSheet sheet, List<ExcelHeaderCell> headers, Collection<T> dataset, String pattern, int contentRow) {
        HSSFCellStyle contentStyle = ExcelStyleUtil.createContentStyle(workbook);
        for (T t : dataset) {
            HSSFRow row = sheet.createRow(contentRow++);
            if (t instanceof Map) {
                setupCellsFromMap(workbook, row, contentStyle, headers, pattern, (Map<String, Object>) t);
            } else {
                setupCellsFromBean(workbook, row, contentStyle, pattern, t);
            }
        }
    }

    private static <T> void setupContentRows(HSSFWorkbook workbook, HSSFSheet sheet, LinkedHashMap<String, String> headers, Collection<T> dataset, String pattern, HSSFCellStyle contentStyle, int contentRow) {
        for (T t : dataset) {
            HSSFRow row = sheet.createRow(contentRow++);
            if (t instanceof Map) {
                setupCellsFromMap(workbook, row, contentStyle, headers, pattern, (Map<String, Object>) t);
            } else {
                setupCellsFromBean(workbook, row, contentStyle, pattern, t);
            }
        }
    }

    private static int setCellValue(HSSFWorkbook workbook, HSSFCell cell, Object value, String pattern, int cellNum, Field field, HSSFRow row) {
        String textValue = null;
        if (value instanceof Integer) {
            int intValue = (Integer) value;
            cell.setCellValue(intValue);
        } else if (value instanceof Float) {
            float fValue = (Float) value;
            cell.setCellValue(fValue);
        } else if (value instanceof Double) {
            double dValue = (Double) value;
            cell.setCellValue(dValue);
        } else if (value instanceof Long) {
            long longValue = (Long) value;
            cell.setCellValue(longValue);
        } else if (value instanceof Boolean) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            textValue = sdf.format(date);
        } else if (value instanceof ZonedDateTime) {
            ZonedDateTime date = (ZonedDateTime) value;
            textValue = date.format(DateTimeFormatter.ofPattern(pattern));
        } else if (value instanceof String[]) {
            String[] strArr = (String[]) value;
            for (int j = 0; j < strArr.length; j++) {
                String str = strArr[j];
                cell.setCellValue(str);
                if (j != strArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else if (value instanceof Double[]) {
            Double[] douArr = (Double[]) value;
            for (int j = 0; j < douArr.length; j++) {
                Double val = douArr[j];
                // 值不为空则set Value
                if (val != null) {
                    cell.setCellValue(val);
                }

                if (j != douArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else {
            // 其它数据类型都当作字符串简单处理
            String empty = StringUtils.EMPTY;
            if (field != null) {
                ExcelCell anno = field.getAnnotation(ExcelCell.class);
                if (anno != null) {
                    empty = anno.defaultValue();
                }
            }
            textValue = value == null ? empty : value.toString();
        }
        if (textValue != null) {
            if (field != null) {
                ExcelCell anno = field.getAnnotation(ExcelCell.class);
                HSSFRichTextString richString = new HSSFRichTextString(textValue);
                cell.setCellValue(richString);
                if (anno.wrap()) {
                    HSSFCellStyle cellStyle = workbook.createCellStyle();
                    cellStyle.setWrapText(true);
                    cell.setCellStyle(cellStyle);
                }
            } else {
                HSSFRichTextString richString = new HSSFRichTextString(textValue);
                cell.setCellValue(richString);
            }
        }
        return cellNum;
    }

    private static <T> void setupCellsFromMap(HSSFWorkbook workbook, HSSFRow row, HSSFCellStyle contentStyle, LinkedHashMap<String, String> headers, String pattern, Map<String, Object> t) {
        @SuppressWarnings("unchecked")
        Map<String, Object> map = t;
        int cellNum = 0;
        //遍历列名
        for (Map.Entry<String, String> entry : headers.entrySet()) {
            Object value = map.get(entry.getKey());
            HSSFCell cell = row.createCell(cellNum);
            cell.setCellStyle(contentStyle);
            cellNum = setCellValue(workbook, cell, value, pattern, cellNum, null, row);
            cellNum++;
        }
    }

    private static <T> void setupCellsFromMap(HSSFWorkbook workbook, HSSFRow row, HSSFCellStyle contentStyle, List<ExcelHeaderCell> headers, String pattern, Map<String, Object> t) {
        Map<String, Object> map = t;
        int cellNum = 0;
        //遍历列名
        for (ExcelHeaderCell excelHeaderCell : headers) {
            if (StringUtils.isNotEmpty(excelHeaderCell.getProperty())) {
                Object value = map.get(excelHeaderCell.getProperty());
                HSSFCell cell = row.createCell(cellNum);
                cell.setCellStyle(contentStyle);
                cellNum = setCellValue(workbook, cell, value, pattern, cellNum, null, row);
                cellNum++;
            }
        }
    }

    private static <T> void setupCellsFromBean(HSSFWorkbook workbook, HSSFRow row, HSSFCellStyle contentStyle, String pattern, T t) {
        try {
            List<FieldForSorting> fields = sortFieldByAnno(t.getClass());
            int cellNum = 0;
            for (FieldForSorting field1 : fields) {
                HSSFCell cell = row.createCell(cellNum);
                cell.setCellStyle(contentStyle);
                Field field = field1.getField();
                field.setAccessible(true);
                Object value = field.get(t);
                cellNum = setCellValue(workbook, cell, value, pattern, cellNum, field, row);
                cellNum++;
            }
        } catch (Exception e) {
            LG.error(e.toString(), e);
        }
    }
}
