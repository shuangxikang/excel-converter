package com.ksx.tools.excel.style;

import org.apache.poi.ss.usermodel.*;

/**
 * Created by ksx on 2019-01-26.
 */
public class RedDataCellStyle implements DataCellStyle {

    /* 数据起始行索引 */
    public int startDataRowNumber = 1;

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();

        //1.0 设置背景色:
        cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());// 设置背景色
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //2.0 设置字体:
        Font font = workbook.createFont();
        font.setFontName("楷体");
        font.setFontHeightInPoints((short) 10);//设置字体大小
        cellStyle.setFont(font);//选择需要用到的字体格式

        return cellStyle;
    }

    @Override
    public int getStartDataRowNumber() {
        return startDataRowNumber;
    }
}
