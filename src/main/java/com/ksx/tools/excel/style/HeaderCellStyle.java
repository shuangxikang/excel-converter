package com.ksx.tools.excel.style;


/**
 * Excel 表头单元格样式，实现本接口即可使用自定义表头样式
 * Created by kangshuangxi on 2016/12/30.
 */
public interface HeaderCellStyle extends CellHelper {

    /**
     * 是否创建表头
     * @return
     */
    boolean isCreateHeader();

    /**
     * 表头行索引
     * @return
     */
    int getHeaderRowNumber();
}
