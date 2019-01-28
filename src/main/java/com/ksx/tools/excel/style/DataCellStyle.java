package com.ksx.tools.excel.style;

/**
 * Excel 表格数据行样式，实现本接口即可实现自定义单元格样式
 * Created by kangshuangxi on 2016/12/30.
 */
public interface DataCellStyle extends CellHelper {

    /**
     * 数据起始行索引
     * @return
     */
    int getStartDataRowNumber();
}
