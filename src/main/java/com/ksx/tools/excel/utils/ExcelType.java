package com.ksx.tools.excel.utils;

import java.util.HashMap;
import java.util.Map;

/**
 * Excel 工作簿类型
 * Created by kangshuangxi on 2016/12/30.
 */
public enum ExcelType {
    /* 老版本Excel Office 97-2004及之前版本兼容使用，文件名以".xls"结尾 */
    xls("application/vnd.ms-excel", ".xls"),
    /* 新版本Excel Office 2004之后版本，文件名以".xlsx"结尾 */
    xlsx("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx");
    ExcelType(String contentType, String filenameExtension) {
        this.contentType = contentType;
        this.filenameExtension = filenameExtension;
    }

    /* HTTP ContentType 类型 */
    private final String contentType;
    /* 文件扩展名 */
    private final String filenameExtension;

    public String getContentType() {
        return contentType;
    }

    public String getFilenameExtension() {
        return filenameExtension;
    }

    /* 类型表（编译通过contentType或filenameExtension获取类型） */
    private static final Map<String, ExcelType> TYPE_MAP = new HashMap();
    static {
        for (ExcelType type : ExcelType.values()) {
            TYPE_MAP.put(type.contentType, type);
            TYPE_MAP.put(type.filenameExtension, type);
        }
    }

    /**
     * 根据请求类型获取 Excel 类型
     * @param contentType HTTP Content Type
     * @return  不存在则返回 null
     */
    public static ExcelType getExcelTypeByContentType(String contentType) {
        return TYPE_MAP.get(contentType);
    }

    /**
     * 根据文件扩展名获取文件类型
     * @param filenameExtension 文件扩展名
     * @return  不存在返回 null
     */
    public static ExcelType getExcelTypeByFilenameExtension(String filenameExtension) {
        return TYPE_MAP.get(filenameExtension);
    }

    /**
     * 根据文件判断 Excel 文件类型
     * @param filename  文件名包含扩展名
     * @return  不存在返回 null
     */
    public static ExcelType getExcelTypeByFilename(String filename) {
        if (filename == null || filename.trim().length() == 0)
            return null;

        for (ExcelType type : ExcelType.values()) {
            if (filename.endsWith(type.filenameExtension))
                return type;
        }

        return null;
    }
}
