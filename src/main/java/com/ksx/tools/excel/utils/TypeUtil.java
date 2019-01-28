package com.ksx.tools.excel.utils;

import java.util.HashMap;
import java.util.Map;

/**
 * 类型判断工具类
 * Created by kangshuangxi on 2016/12/30.
 */
public class TypeUtil {

    /* 数值类型列表 */
    private static final Map<String, String> NUMBER_TYPES = new HashMap(){
        {
            /* 数值包装类型 */
            put(Short.class.getName(), "short");
            put(Integer.class.getName(), "int");
            put(Long.class.getName(), "long");
            put(Float.class.getName(), "float");
            put(Double.class.getName(), "double");

            /* 数值基本类型 */
            put("short", "short");
            put("int", "int");
            put("long", "long");
            put("float", "float");
            put("double", "double");
        }
    };

    /**
     * 根据类型名称判断
     * @param typeName  类型名称
     * @return
     */
    public static boolean isNumber(String typeName) {
        return NUMBER_TYPES.get(typeName) != null;
    }

    /**
     * 是否为布尔类型
     * @param typeName  类型名称
     * @return
     */
    public static boolean isBoolean(String typeName) {
        return typeName != null && (typeName.equals("boolean") || typeName.equals(Boolean.TYPE.getName()));
    }

    /**
     * 是否为字符类型
     * @param typeName  类型名称
     * @return
     */
    public static boolean isChar(String typeName) {
        return typeName != null && (typeName.equals("char") || typeName.equals(Character.TYPE.getName()));
    }
}
