package com.ksx.tools.excel.utils;

import com.ksx.tools.excel.Excel;
import com.ksx.tools.excel.annotation.ExcelColumn;
import com.ksx.tools.excel.annotation.ExcelColumnSplitCell;
import com.ksx.tools.excel.style.CellSplitStyle;
import com.ksx.tools.excel.style.DefaultCellSplitStyle;

import java.lang.reflect.Field;
import java.util.*;

/**
 * 反射工具类
 * Created by kangshuangxi on 2016/12/30.
 */
public class ReflectUtil {

    /**
     * 获取所有加了ExcelCell注解的字段
     * @param clazz     类型
     * @return
     */
    public static List<Field> getDeclaredFields(Class clazz) {
        if (clazz == null)
            return null;

        List<Field> fields = new ArrayList();
        return getDeclaredFields(clazz, fields);
    }

    /**
     * 获取所有加了ExcelCell注解的字段
     * @param clazz     类型
     * @param fields    输出字段列表
     * @return
     */
    public static List<Field> getDeclaredFields(Class clazz, List<Field> fields) {
        //1.0 如果有父类得到父类字段
        if (clazz.getSuperclass() != null && !clazz.getSuperclass().getTypeName().endsWith("Object"))
            getDeclaredFields(clazz.getSuperclass(), fields);

        //2.0 得到所有加ExcelCell的fields
        Field[] declaredFields = clazz.getDeclaredFields();
        for (Field field : declaredFields) {
            if (field.isAnnotationPresent(ExcelColumn.class))
                fields.add(field);
        }

        return fields;
    }

    /**
     * 获取字段表
     * @param clazz
     * @return
     */
    public static Map<Integer, Field> getFieldMap(Class clazz) {
        //1.0 获取有注解的字段列表
        List<Field> allFields = getDeclaredFields(clazz);
        if (allFields == null || allFields.size() == 0)
            return null;

        //2.0 将field对应到excel列索引：Map<key=columnIndex, value=field>
        int columnIndex;
        String columnSta;
        Map<Integer, Field> fieldsMap = new HashMap();
        for (Field field : allFields) {
            columnSta = getColumn(field);

            //2.1 转换excel字母序到下标索引
            columnIndex = Excel.getExcelColumnIndex(columnSta);
            //2.2 设置类的私有字段属性可访问，方便后面反射设置属性值
            field.setAccessible(true);
            fieldsMap.put(columnIndex, field);
        }

        return fieldsMap;
    }

    public static Map<Integer, Field> getAllHeaderFiled(Map<Integer, Field> fieldMap) {
        Map<Integer, Field> headerFields = new HashMap();
        headerFields.putAll(fieldMap);
        CellSplitStyle cellSplitStyle = getSplitField(fieldMap);
        if (cellSplitStyle != null)
            headerFields.putAll(cellSplitStyle.getItemFields());

        return headerFields;
    }

    private static String getColumn(Field field) {
        if (field == null)
            return null;

        String columnSta = null;
        ExcelColumn attr = field.getAnnotation(ExcelColumn.class);
        if (attr != null)
            return attr.column();

        return columnSta;
    }

    /**
     * 获取需要拆分单元格字段
     * @param fieldMap
     * @return
     */
    public static CellSplitStyle getSplitField(Map<Integer, Field> fieldMap) {
        if (fieldMap == null)
            return null;

        for (Field field : fieldMap.values()) {
            ExcelColumnSplitCell splitCell = field.getAnnotation(ExcelColumnSplitCell.class);
            if (splitCell != null) {
                int beginColumn = Excel.getExcelColumnIndex(splitCell.beginColumn());
                int endColumn = Excel.getExcelColumnIndex(splitCell.endColumn());
                Map<Integer, Field> fields = getFieldMap(splitCell.itemClass());
                return new DefaultCellSplitStyle(field, fields, beginColumn, endColumn);
            }
        }

        return null;
    }

    public static Object getFiledValue(Field field, Object data) {
        Object fieldValue = null;
        try {
            fieldValue = field.get(data);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        return fieldValue;
    }

    public static int getCollectionSize(Object data) {
        if (data == null)
            return 0;

        if (data instanceof Collection)
            return ((Collection) data).size();

        if (data instanceof Map)
            return ((Map) data).size();

        return 0;
    }
}
