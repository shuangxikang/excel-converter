package com.ksx.tools.excel;

import com.ksx.tools.excel.annotation.ExcelColumn;
import com.ksx.tools.excel.format.DataFormat;
import com.ksx.tools.excel.utils.ExcelType;
import com.ksx.tools.excel.utils.ReflectUtil;
import com.ksx.tools.excel.utils.TypeUtil;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Excel导入
 * Created by kangshuangxi on 2016/12/30.
 */
public class ExcelImport<T> {
    private static final Logger log = LoggerFactory.getLogger(ExcelImport.class);

    /* 导入对象类型 */
    private Class<T> clazz;
    public ExcelImport(Class<T> clazz) {
        this.clazz = clazz;
    }

    /**
     * Excel 导入
     * @param input     文件输入流
     * @param type      Excel类型
     * @return
     */
    public List<T> importExcel(InputStream input, ExcelType type) {
        return importExcel(input, type, "");
    }

    /**
     * Excel 导入
     * @param input     文件输入流
     * @param type      Excel类型
     * @param sheetName 需要解析的工作表名称
     * @return
     */
    public List<T> importExcel(InputStream input, ExcelType type, String sheetName) {
        try {
            //1.0 创建工作表
            Workbook workbook = Excel.createWorkbook(input, type);

            //2.0 获取指定工作簿（如不存在默认获取第一个工作簿）
            Sheet sheet;
            if (sheetName == null || sheetName.trim().equals("")) {
                sheet = workbook.getSheetAt(0);
            } else {
                sheet = workbook.getSheet(sheetName);
            }

            return readExcelSheet(sheet);
        } catch (Exception e) {
            log.error("表格导入异常", e);
        }

        return null;
    }

    /**
     * 读取工作表
     * @param sheet
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public List<T> readExcelSheet(Sheet sheet) throws IllegalAccessException, InstantiationException {
        if (sheet == null)
            return null;

        List<T> list = null;
        int rows = sheet.getLastRowNum();
        if (rows > 0) {
            list = new ArrayList();
            Map<Integer, Field> fieldsMap = ReflectUtil.getFieldMap(clazz);
            int startDataRowNumber = getStartDataRowNumber(fieldsMap);
            for (int i = startDataRowNumber; i <= rows; i++) {
                Row row = sheet.getRow(i);
                T entity = readExcelRow(fieldsMap, row);

                if (entity != null)
                    list.add(entity);
            }
        }
        return list;
    }

    /**
     * 读取一行数据
     * @param fieldsMap
     * @param row
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    private T readExcelRow(Map<Integer, Field> fieldsMap, Row row) throws IllegalAccessException, InstantiationException {
        if (fieldsMap == null || fieldsMap.size() == 0 || row == null)
            return null;

        T rowItem = clazz.newInstance();
        for (Map.Entry entry : fieldsMap.entrySet()) {
            Cell cell = row.getCell((int)entry.getKey());
            Field field = (Field)entry.getValue();
            if (cell == null || field == null)
                continue;

            setFieldValue(rowItem, field, cell);
        }

        return rowItem;
    }

    /**
     * 设置对象属性
     * @param rowItem
     * @param field
     * @param cell
     * @throws IllegalAccessException
     */
    private void setFieldValue(T rowItem, Field field, Cell cell) throws IllegalAccessException {
        Class fieldType = field.getType();
        if (fieldType.isAssignableFrom(String.class)) {
            CellType cellType = cell.getCellTypeEnum();
            String value;
            if (cellType == CellType.NUMERIC) {
                Double doubleValue = cell.getNumericCellValue();
                if (doubleValue > doubleValue.longValue()) {
                    value = String.valueOf(doubleValue);
                } else {
                    value = String.valueOf(Double.valueOf(cell.getNumericCellValue()).longValue());
                }
            } else {
                //强制设置cell类型, 否则有可能读取失败
                cell.setCellType(CellType.STRING);
                value = cell.getStringCellValue();
            }
            field.set(rowItem, value);
        } else if (TypeUtil.isNumber(fieldType.getName())) {
            setNumberValue(rowItem, field, cell);
        } else if (TypeUtil.isBoolean(fieldType.getName())) {
            field.set(rowItem, cell.getBooleanCellValue());
        } else if (TypeUtil.isChar(fieldType.getName())) {
            if (cell.getStringCellValue() != null)
                field.set(rowItem, cell.getStringCellValue().charAt(0));
        } else if (fieldType.isAssignableFrom(Date.class)) {
            field.set(rowItem, cell.getDateCellValue());
        } else if (fieldType.isAssignableFrom(BigDecimal.class)) {
            field.set(rowItem, new BigDecimal(cell.getNumericCellValue()));
        } else {
            setOtherValue(rowItem, field, cell);
        }
    }

    /**
     * 设置数值属性
     * @param rowItem
     * @param field
     * @param cell
     * @throws IllegalAccessException
     */
    private void setNumberValue(T rowItem, Field field, Cell cell) throws IllegalAccessException {
        Double cellValue = Double.valueOf(cell.getNumericCellValue());
        String typeName = field.getType().getName();
        switch (typeName) {
            case "short":
            case "java.lang.Short":
                field.set(rowItem, cellValue.shortValue());
                break;
            case "int":
            case "java.lang.Integer":
                field.set(rowItem, cellValue.intValue());
                break;
            case "long":
            case "java.lang.Long":
                field.set(rowItem, cellValue.longValue());
                break;
            case "float":
            case "java.lang.Float":
                field.set(rowItem, cellValue.floatValue());
                break;
            case "double":
            case "java.lang.Double":
                field.set(rowItem, cellValue.doubleValue());
                break;
        }
    }

    /**
     * 获取数据起始行设置
     * @param fieldsMap
     * @return
     */
    private int getStartDataRowNumber(Map<Integer, Field> fieldsMap) {
        int num = 1;
        ExcelColumn excelColumn;
        for (Field field : fieldsMap.values()) {
            excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn != null && excelColumn.dataCellStyle() != null) {
                try {
                    num = excelColumn.dataCellStyle().getDeclaredConstructor().newInstance().getStartDataRowNumber();
                } catch (InstantiationException e) {
                    log.error("表格导入异常", e);
                } catch (IllegalAccessException e) {
                    log.error("表格导入异常", e);
                } catch (NoSuchMethodException e) {
                    log.error("表格导入异常", e);
                } catch (InvocationTargetException e) {
                    log.error("表格导入异常", e);
                }
            }
        }

        return num;
    }

    /**
     * 其他类型转换
     * @param rowItem
     * @param field
     * @param cell
     */
    public void setOtherValue(T rowItem, Field field, Cell cell) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        Class<DataFormat> dataFormat = (Class<DataFormat>) excelColumn.dataFormat();

        //TODO 暂无实现，后续完善
    }
}
