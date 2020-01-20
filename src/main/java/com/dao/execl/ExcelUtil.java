package com.dao.execl;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ExcelUtil {

    private static final Integer SHEET_MAX_DATA_COUNT = 5000;


    /**
     * excel导出，默认xls
     *
     * @param dataList
     * @param <T>
     * @return
     */
    public static <T> Workbook excelExport(List<T> dataList) {
        return excelExport(dataList, ExcelType.XLS);
    }

    /**
     * excel导出，可以指定xls或者xlsx
     *
     * @param dataList
     * @param excelType
     * @param <T>
     * @return
     */
    public static <T> Workbook excelExport(List<T> dataList, ExcelType excelType) {
        if (dataList == null || dataList.isEmpty()) {
            return null;
        }
        Class<?> clazz = dataList.get(0).getClass();
        List<ExcelField> excelFieldList = resolveExcelAnno(clazz);
        if (excelFieldList.isEmpty()) {
            return null;
        }
        Workbook workbook;
        if (excelType.equals(ExcelType.XLS)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }

        int dataSize = dataList.size();
        int sheetCount = (dataSize / SHEET_MAX_DATA_COUNT) + 1;
        for (int i = 0; i < sheetCount; i++) {
            int startIndex = i * SHEET_MAX_DATA_COUNT;
            int tailIndex = startIndex + SHEET_MAX_DATA_COUNT;
            tailIndex = tailIndex >= (dataSize - 1) ? dataSize - 1 : tailIndex;
            createSheet(workbook, excelFieldList, dataList, startIndex, tailIndex);
        }
        return workbook;
    }

    /**
     * excel导入解析
     *
     * @param is
     * @param clazz
     * @param <T>
     * @return
     * @throws IOException
     */
    public static <T> List<T> excelImport(InputStream is, Class<T> clazz) throws IOException {
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(is);
            return excelImport(workbook, clazz);
        } finally {
            if (workbook != null) {
                workbook.close();
            }
        }

    }

    /**
     * excel导入解析
     *
     * @param workbook
     * @param clazz
     * @param <T>
     * @return
     */
    public static <T> List<T> excelImport(Workbook workbook, Class<T> clazz) {
        List<ExcelField> excelFieldList = resolveExcelAnno(clazz);
        if (excelFieldList.isEmpty()) {
            return Collections.emptyList();
        }
        int sheetCount = workbook.getNumberOfSheets();
        if (sheetCount <= 0) {
            return Collections.emptyList();
        }
        List<T> dataList = new ArrayList<>();
        for (int i = 0; i < sheetCount; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            Map<Integer, ExcelField> cellFieldMap = getCellIndexAndFieldMap(sheet, excelFieldList);
            if (cellFieldMap == null) {
                continue;
            }
            int rowCount = sheet.getPhysicalNumberOfRows();
            if (rowCount <= 1) {
                continue;
            }
            Iterator<Row> rowIterator = sheet.rowIterator();
            rowIterator.next();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                T obj = resolveRow(row, cellFieldMap, clazz);
                if (obj == null) {
                    continue;
                }
                dataList.add(obj);
            }
        }
        return dataList;
    }

    /**
     * 设置每一页的内容
     *
     * @param workbook
     * @param excelFieldList
     * @param dataList
     * @param startIndex
     * @param tailIndex
     * @param <T>
     */
    private static <T> void createSheet(Workbook workbook, List<ExcelField> excelFieldList, List<T> dataList, int startIndex, int tailIndex) {
        Sheet sheet = workbook.createSheet();
        CellStyle cellStyle = getHeadStyle(workbook);
        Map<String, Integer> cellIndexFieldMap = createHeadAndGetCellIndexFiledMap(sheet, cellStyle, excelFieldList);

        Map<String, ExcelField> fieldNameMap = new HashMap<>(excelFieldList.size());
        for (ExcelField excelField : excelFieldList) {
            fieldNameMap.put(excelField.getFieldName(), excelField);
        }
        int rowIndex = 1;
        for (int i = startIndex; i <= tailIndex; i++) {
            Row row = sheet.createRow(rowIndex++);
            T t = dataList.get(i);
            Field[] fields = t.getClass().getDeclaredFields();
            for (Field field : fields) {
                Integer cellIndex = cellIndexFieldMap.get(field.getName());
                if (cellIndex != null) {
                    Cell cell = row.createCell(cellIndex);
                    Object fieldValue = null;
                    try {
                        PropertyDescriptor descriptor = new PropertyDescriptor(field.getName(), t.getClass());
                        Method readMethod = descriptor.getReadMethod();
                        fieldValue = readMethod.invoke(t);
                    } catch (Exception e) {
                        throw new ExcelException("获取对象的字段值错误,class :" + t.getClass() + "; fieldName: " + field.getName());
                    }
                    if (fieldValue != null) {
                        setCellValue(cell, fieldValue, fieldNameMap.get(field.getName()));
                    } else {
                        cell.setBlank();
                    }

                }
            }
        }
    }

    /**
     * 获取表头样式
     *
     * @param workbook
     * @return
     */
    private static CellStyle getHeadStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        //水平居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //设置背景色
//        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
//        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font font = workbook.createFont();
        //设置粗体
        font.setBold(true);
        //设置字体颜色
        font.setColor(Font.COLOR_NORMAL);
        //设置字体高度
        font.setFontHeightInPoints((short) 10);
        cellStyle.setFont(font);
        return cellStyle;
    }


    /**
     * 设置列的值
     *
     * @param cell
     * @param fieldValue
     * @param excelField
     */
    private static void setCellValue(Cell cell, Object fieldValue, ExcelField excelField) {
        Class fieldType = excelField.getFieldType();
        if (fieldType == int.class || fieldType == Integer.class) {
            cell.setCellValue((Integer) fieldValue);
        } else if (fieldType == double.class || fieldType == Double.class) {
            cell.setCellValue((Double) fieldValue);
        } else if (fieldType == boolean.class || fieldType == Boolean.class) {
            cell.setCellValue((Boolean) fieldValue);
        } else if (fieldType == BigDecimal.class) {
            cell.setCellValue(((BigDecimal) fieldValue).doubleValue());
        } else if (fieldType == Date.class) {
            DateFormat dateFormat = new SimpleDateFormat(excelField.getFormat());
            dateFormat.setTimeZone(TimeZone.getTimeZone(excelField.getTimezone()));
            cell.setCellValue(dateFormat.format((Date) fieldValue));
        } else if (fieldType == LocalDateTime.class) {
            LocalDateTime time = (LocalDateTime) fieldValue;
            cell.setCellValue(time.format(DateTimeFormatter.ofPattern(excelField.getFormat())));
        } else if (fieldType == String.class) {
            cell.setCellValue((String) fieldValue);
        } else {
            throw new ExcelException("不支持的类型转换, fieldType: " + fieldType.getName());
        }
    }

    /**
     * 创建表头，并获取列号和字段名称的关系
     *
     * @param sheet
     * @param cellStyle
     * @param excelFieldList
     * @return
     */
    private static Map<String, Integer> createHeadAndGetCellIndexFiledMap(Sheet sheet, CellStyle cellStyle, List<ExcelField> excelFieldList) {
        Row row = sheet.createRow(0);
        row.setHeight((short) (30 * 20));
        int cellIndex = 0;
        Map<String, Integer> cellIndexFieldMap = new HashMap<>(excelFieldList.size());
        excelFieldList.sort(Comparator.comparing(ExcelField::getCellIndexSort));
        for (ExcelField excelField : excelFieldList) {
            sheet.setColumnWidth(cellIndex,20*256);
            Cell cell = row.createCell(cellIndex);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(excelField.getCellName());
            cellIndexFieldMap.put(excelField.getFieldName(), cellIndex);
            cellIndex++;
        }

        return cellIndexFieldMap;
    }


    /**
     * 解析行数据
     *
     * @param row
     * @param cellFieldMap
     * @param clazz
     * @param <T>
     * @return
     */
    private static <T> T resolveRow(Row row, Map<Integer, ExcelField> cellFieldMap, Class<T> clazz) {
        try {
            if (row == null) {
                return null;
            }
            int cellCount = row.getLastCellNum();
            if (cellCount <= 0) {
                return null;
            }
            T obj = clazz.newInstance();
            for (int cellIndex = 0; cellIndex < cellCount; cellIndex++) {
                ExcelField excelField = cellFieldMap.get(cellIndex);
                if (excelField != null) {
                    Cell cell = row.getCell(cellIndex);
                    Object fieldValue = getFieldValue(cell, excelField);
                    PropertyDescriptor descriptor = new PropertyDescriptor(excelField.getFieldName(), clazz);
                    Method method = descriptor.getWriteMethod();
                    method.invoke(obj, fieldValue);
                }
            }
            return obj;
        } catch (Exception e) {
            throw new ExcelException("解析excel行失败", e);
        }


    }

    /**
     * 获取列和字段的对应关系
     *
     * @param sheet
     * @param excelFieldList
     * @return
     */
    private static <T> Map<Integer, ExcelField> getCellIndexAndFieldMap(Sheet sheet, List<ExcelField> excelFieldList) {
        Row row = sheet.getRow(0);
        if (row == null) {
            return null;
        }
        int cellCount = row.getPhysicalNumberOfCells();
        if (cellCount <= 0) {
            return null;
        }
        Map<String, ExcelField> excelFieldMap = new HashMap<>(excelFieldList.size());
        for (ExcelField excelField : excelFieldList) {
            excelFieldMap.put(excelField.getCellName(), excelField);
        }
        Map<Integer, ExcelField> cellIndexFieldMap = new HashMap<>(excelFieldMap.size());
        for (int i = 0; i < cellCount; i++) {
            Cell cell = row.getCell(i);
            if (!cell.getCellType().equals(CellType.STRING)) {
                throw new ExcelException("第一行必须是字符类型");
            }
            String value = cell.getStringCellValue();
            ExcelField excelField = excelFieldMap.get(value);
            if (excelField != null) {
                cellIndexFieldMap.put(i, excelField);
            }

        }
        return cellIndexFieldMap;
    }

    /**
     * 获取列中数据，并转换为字段对应的值
     *
     * @param cell
     * @param excelField
     * @return
     */
    private static Object getFieldValue(Cell cell, ExcelField excelField) {
        if (cell == null) {
            return null;
        }
        String cellValue = getCellValue(cell);
        Class fieldType = excelField.getFieldType();
        Object realCellValue = null;
        if (fieldType == int.class || fieldType == Integer.class) {
            realCellValue = Double.valueOf(cellValue).intValue();
        } else if (fieldType == double.class || fieldType == Double.class) {
            realCellValue = Double.parseDouble(cellValue);
        } else if (fieldType == boolean.class || fieldType == Boolean.class) {
            realCellValue = Boolean.parseBoolean(cellValue);
        } else if (fieldType == BigDecimal.class) {
            realCellValue = BigDecimal.valueOf(Double.parseDouble(cellValue));
        } else if (fieldType == LocalDateTime.class) {
            realCellValue = LocalDateTime.parse(cellValue, DateTimeFormatter.ofPattern(excelField.getFormat()));
        } else if (fieldType == Date.class) {
            LocalDateTime time = LocalDateTime.parse(cellValue, DateTimeFormatter.ofPattern(excelField.getFormat()));
            realCellValue = Date.from(time.atZone(ZoneId.of(excelField.getTimezone())).toInstant());
        } else if (fieldType == String.class) {
            realCellValue = cellValue;
        } else {
            throw new ExcelException("不支持的类型转换, fieldType: " + fieldType.getName());
        }
        return realCellValue;
    }

    /**
     * 获取列中数据
     *
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    /**
     * 解析类上的ExcelCell注解
     *
     * @param clazz
     * @return
     */
    private static <T> List<ExcelField> resolveExcelAnno(Class<T> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        List<ExcelField> excelFieldList = new ArrayList<>(fields.length);
        if (fields.length <= 0) {
            return excelFieldList;
        }
        for (Field field : fields) {
            ExcelField excelField = new ExcelField();
            excelField.setFieldName(field.getName());
            excelField.setFieldType(field.getType());
            excelField.setCellName(field.getName());
            excelField.setCellIndexSort(0);
            ExcelCell annotation = field.getAnnotation(ExcelCell.class);
            if (annotation != null) {
                excelField.setCellName(annotation.value());
                excelField.setFormat(annotation.format());
                excelField.setTimezone(annotation.timezone());
                excelField.setCellIndexSort(annotation.cellIndexSort());
            }
            excelFieldList.add(excelField);
        }
        validateExcelFieldList(excelFieldList);
        return excelFieldList;
    }

    private static void validateExcelFieldList(List<ExcelField> excelFieldList) {
        for (ExcelField excelField : excelFieldList) {
            if (excelField.getCellIndexSort() < 0) {
                throw new ExcelException("cellIndex不能小于0");
            }
        }
        long cellNameSize = excelFieldList.stream().map(ExcelField::getCellName).distinct().count();
        if (excelFieldList.size() != cellNameSize) {
            throw new ExcelException("存在重复的列名");
        }
    }

}
