package com.ksx.tools.excel.style;

import com.ksx.tools.excel.utils.ExcelType;

/**
 * 导出表格样式设置
 * Created by kangshuangxi on 2016/12/30.
 */
public interface ExcelStyle {

    /**
     * 设置表格类型
     * @return
     */
    ExcelType getExcelType();

    /**
     * 表头行索引
     * @return
     */
    int getStartHeaderRowNumber();

    /**
     * 数据起始行索引
     * @return
     */
    int getStartDataRowNumber();
}
