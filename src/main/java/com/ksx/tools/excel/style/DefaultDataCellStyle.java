package com.ksx.tools.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 数据Cell样式
 * Created by kangshuangxi on 2016/12/31.
 */
public class DefaultDataCellStyle implements DataCellStyle {

    /* 数据起始行索引 */
    public int startDataRowNumber = 1;

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        //1.0 设置字体:
        Font font = workbook.createFont();
        font.setFontName("楷体");
        font.setFontHeightInPoints((short) 10);//设置字体大小
//        font.setBold(true);//粗体显示
        cellStyle.setFont(font);//选择需要用到的字体格式

        //设置自动换行:
//        cellStyle.setWrapText(true);//设置自动换行
        return cellStyle;
    }

    @Override
    public int getStartDataRowNumber() {
        return startDataRowNumber;
    }
}
