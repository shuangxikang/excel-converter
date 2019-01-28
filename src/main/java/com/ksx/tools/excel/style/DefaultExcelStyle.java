package com.ksx.tools.excel.style;

import com.ksx.tools.excel.utils.ExcelType;

/**
 * Excel 表格样式
 * Created by kangshuangxi on 2016/12/31.
 */
public class DefaultExcelStyle implements ExcelStyle {

    /* 表格类型 */
    private final ExcelType excelType;
    /* 表头开始行 */
    private int startHeaderRowNumber = 0;
    /* 数据单元格开始行 */
    private int startDataRowNumber = 1;

    public DefaultExcelStyle(ExcelType excelType) {
        if (excelType == null)
            throw new ExceptionInInitializerError("excelType 不能为空");

        this.excelType = excelType;
    }

    @Override
    public ExcelType getExcelType() {
        return excelType;
    }

    @Override
    public int getStartHeaderRowNumber() {
        return startHeaderRowNumber;
    }

    @Override
    public int getStartDataRowNumber() {
        return startDataRowNumber;
    }
}
