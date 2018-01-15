package com.sargeraswang.util.ExcelUtil;

import org.apache.commons.beanutils.BeanComparator;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.ComparatorUtils;
import org.apache.commons.collections.comparators.ComparableComparator;
import org.apache.commons.collections.comparators.ComparatorChain;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * The <code>ExcelUtil</code> 与 {@link ExcelCell}搭配使用
 *
 * @author sargeras.wang
 * @version 1.0, Created at 2013年9月14日
 */
public class ExcelImportUtil extends ExcelUtil {

    private static Logger LG = LoggerFactory.getLogger(ExcelImportUtil.class);

    /**
     * 用来验证excel与Vo中的类型是否一致 <br> Map<栏位类型,只能是哪些Cell类型>
     */
    private static Map<Class<?>, CellType[]> validateMap = new HashMap<>();

    static {
        validateMap.put(String[].class, new CellType[]{CellType.STRING});
        validateMap.put(Double[].class, new CellType[]{CellType.NUMERIC});
        validateMap.put(String.class, new CellType[]{CellType.STRING});
        validateMap.put(Double.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Date.class, new CellType[]{CellType.NUMERIC, CellType.STRING});
        validateMap.put(Integer.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Float.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Long.class, new CellType[]{CellType.NUMERIC});
        validateMap.put(Boolean.class, new CellType[]{CellType.BOOLEAN});
    }

    public static <T> Collection<T> importExcel(File excelFile, Class<T> clazz, Integer ignoreRow, ExcelLogs logs, Integer... arrayCount) {
        FileInputStream inputStream = null;
        try {
            if (excelFile == null || !excelFile.exists()) {
                throw new RuntimeException("Cannot find the excel file");
            }
            if (ignoreRow == null) {
                ignoreRow = 0;
            }

            inputStream = new FileInputStream(excelFile);
            Workbook workBook = WorkbookFactory.create(inputStream);
            List<T> list = new ArrayList<>();
            List<ExcelLog> logList = new ArrayList<>();
            Sheet sheet = workBook.getSheetAt(0);

            for (int i = ignoreRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                // 整行都空，就跳过
                boolean allRowIsNull = true;
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Object cellValue = getCellValue(cellIterator.next());
                    if (cellValue != null) {
                        allRowIsNull = false;
                        break;
                    }
                }
                if (allRowIsNull) {
                    LG.warn("Excel row " + row.getRowNum() + " all row value is null!");
                    continue;
                }

                StringBuilder log = new StringBuilder();
                if (clazz == Map.class) {
                    Map<Integer, Object> map = new HashMap<>();
                    Integer index = 0;
                    cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        cell.setCellType(CellType.STRING);
                        String value = cell.getStringCellValue();
                        map.put(index++, value);
                    }
                    list.add((T) map);
                } else {
                    T t = clazz.newInstance();
                    int arrayIndex = 0;// 标识当前第几个数组了
                    int cellIndex = 0;// 标识当前读到这一行的第几个cell了
                    List<FieldForSorting> fields = sortFieldByAnno(clazz);
                    for (FieldForSorting ffs : fields) {
                        Field field = ffs.getField();
                        field.setAccessible(true);
                        if (field.getType().isArray()) {
                            Integer count = arrayCount[arrayIndex];
                            Object[] value;
                            if (field.getType().equals(String[].class)) {
                                value = new String[count];
                            } else {
                                // 目前只支持String[]和Double[]
                                value = new Double[count];
                            }
                            for (int j = 0; j < count; j++) {
                                Cell cell = row.getCell(cellIndex);
                                if (field.getType().equals(String[].class)) {
                                    cell.setCellType(CellType.STRING);
                                }
                                String errMsg = validateCell(cell, field, cellIndex);
                                if (StringUtils.isBlank(errMsg)) {
                                    value[j] = getCellValue(cell);
                                } else {
                                    log.append(errMsg).append(";");
                                    logs.setHasError(true);
                                }
                                cellIndex++;
                            }
                            field.set(t, value);
                            arrayIndex++;
                        } else {
                            Cell cell = row.getCell(cellIndex);
                            String errMsg = validateCell(cell, field, cellIndex);
                            if (StringUtils.isBlank(errMsg)) {
                                ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
                                Object value = getCellValue(cell);
                                // 处理特殊情况,Excel中的String,转换成Bean的Date
                                if (value != null) {
                                    if (field.getType().equals(Date.class) && cell.getCellTypeEnum() == CellType.STRING && StringUtils.isNotEmpty(cell.getStringCellValue())) {
                                        try {
                                            value = new SimpleDateFormat(excelCell.dateFormat()).parse(value.toString());
                                        } catch (ParseException e) {
                                            errMsg = MessageFormat.format("the cell [{0}] can not be converted to a date ", CellReference.convertNumToColString(cell.getColumnIndex()));
                                        }
                                    } else if (field.getType().equals(Integer.class)) {
                                        value = ((Double) value).intValue();
                                    }
                                }
                                // 处理特殊情况,excel的value为String,且bean中为其他,且defaultValue不为空,那就=defaultValue
                                if (value instanceof String && !field.getType().equals(String.class) && StringUtils.isNotBlank(excelCell.defaultValue())) {
                                    value = excelCell.defaultValue();
                                }
                                field.set(t, value);
                            }
                            if (StringUtils.isNotBlank(errMsg)) {
                                log.append(errMsg).append(";");
                                logs.setHasError(true);
                            }
                            cellIndex++;
                        }
                    }
                    list.add(t);
                    logList.add(new ExcelLog(t, log.toString(), row.getRowNum() + 1));
                }
            }
            logs.setLogList(logList);
            return list;
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
    }

    /**
     * 驗證Cell類型是否正確
     *
     * @param cell    cell單元格
     * @param field   欄位
     * @param cellNum 第幾個欄位,用於errMsg
     */
    private static String validateCell(Cell cell, Field field, int cellNum) {
        String columnName = CellReference.convertNumToColString(cellNum);
        String result = null;
        CellType[] cellTypeArr = validateMap.get(field.getType());
        if (cellTypeArr == null) {
            result = MessageFormat.format("Unsupported type [{0}]", field.getType().getSimpleName());
            return result;
        }
        ExcelCell annoCell = field.getAnnotation(ExcelCell.class);
        if (cell == null || (cell.getCellTypeEnum() == CellType.STRING && StringUtils.isBlank(cell.getStringCellValue()))) {
            if (annoCell != null && !annoCell.valid().allowNull()) {
                result = MessageFormat.format("the cell [{0}] can not null", columnName);
            }
            ;
        } else if (cell.getCellTypeEnum() == CellType.BLANK && annoCell.valid().allowNull()) {
            return null;
        } else {
            List<CellType> cellTypes = Arrays.asList(cellTypeArr);

            // 如果類型不在指定範圍內,並且沒有默認值
            if (!(cellTypes.contains(cell.getCellTypeEnum()))
                    || StringUtils.isNotBlank(annoCell.defaultValue())
                    && cell.getCellTypeEnum() == CellType.STRING) {
                StringBuilder strType = new StringBuilder();
                for (int i = 0; i < cellTypes.size(); i++) {
                    CellType cellType = cellTypes.get(i);
                    strType.append(getCellTypeByInt(cellType));
                    if (i != cellTypes.size() - 1) {
                        strType.append(",");
                    }
                }
                result = MessageFormat.format("the cell [{0}] type must [{1}]", columnName, strType.toString());
            } else {
                // 类型符合验证,但值不在要求范围内的
                // String in
                if (annoCell.valid().in().length != 0 && cell.getCellTypeEnum() == CellType.STRING) {
                    String[] in = annoCell.valid().in();
                    String cellValue = cell.getStringCellValue();
                    boolean isIn = false;
                    for (String str : in) {
                        if (str.equals(cellValue)) {
                            isIn = true;
                        }
                    }
                    if (!isIn) {
                        result = MessageFormat.format("the cell [{0}] value must in {1}", columnName, in);
                    }
                }
                // 数字型
                if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                    double cellValue = cell.getNumericCellValue();
                    // 小于
                    if (!Double.isNaN(annoCell.valid().lt())) {
                        if (!(cellValue < annoCell.valid().lt())) {
                            result = MessageFormat.format("the cell [{0}] value must less than [{1}]", columnName, annoCell.valid().lt());
                        }
                    }
                    // 大于
                    if (!Double.isNaN(annoCell.valid().gt())) {
                        if (!(cellValue > annoCell.valid().gt())) {
                            result = MessageFormat.format("the cell [{0}] value must greater than [{1}]", columnName, annoCell.valid().gt());
                        }
                    }
                    // 小于等于
                    if (!Double.isNaN(annoCell.valid().le())) {
                        if (!(cellValue <= annoCell.valid().le())) {
                            result = MessageFormat.format("the cell [{0}] value must less than or equal [{1}]", columnName, annoCell.valid().le());
                        }
                    }
                    // 大于等于
                    if (!Double.isNaN(annoCell.valid().ge())) {
                        if (!(cellValue >= annoCell.valid().ge())) {
                            result = MessageFormat.format("the cell [{0}] value must greater than or equal [{1}]", columnName, annoCell.valid().ge());
                        }
                    }
                }
            }
        }
        return result;
    }

    /**
     * 获取cell类型的文字描述
     *
     * @param cellType <pre>
     *                     CellType.BLANK
     *                     CellType.BOOLEAN
     *                     CellType.ERROR
     *                     CellType.FORMULA
     *                     CellType.NUMERIC
     *                     CellType.STRING
     *                 </pre>
     */
    private static String getCellTypeByInt(CellType cellType) {
        if (cellType == CellType.BLANK) {
            return "Null type";
        } else if (cellType == CellType.BOOLEAN) {
            return "Boolean type";
        } else if (cellType == CellType.ERROR) {
            return "Error type";
        } else if (cellType == CellType.FORMULA) {
            return "Formula type";
        } else if (cellType == CellType.NUMERIC) {
            return "Numeric type";
        } else if (cellType == CellType.STRING) {
            return "String type";
        } else {
            return "Unknown type";
        }
    }

    /**
     * 获取单元格值
     */
    private static Object getCellValue(Cell cell) {
        if (cell == null || (cell.getCellTypeEnum() == CellType.STRING && StringUtils.isBlank(cell.getStringCellValue()))) {
            return null;
        }
        CellType cellType = cell.getCellTypeEnum();
        if (cellType == CellType.BLANK) {
            return null;
        } else if (cellType == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        } else if (cellType == CellType.ERROR) {
            return cell.getErrorCellValue();
        } else if (cellType == CellType.FORMULA) {
            try {
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            } catch (IllegalStateException e) {
                return cell.getRichStringCellValue();
            }
        } else if (cellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            } else {
                return cell.getNumericCellValue();
            }
        } else if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else {
            return null;
        }
    }

}
