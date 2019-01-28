package com.ksx.tools.excel.style;

import org.apache.poi.ss.usermodel.*;

/**
 * 表头样式：黄色背景 + 黑色文本
 * Created by kangshuangxi on 2017/5/22.
 */
public class YellowHeaderCellStyle implements HeaderCellStyle {

    /* 是否创建表头 */
    public boolean isCreateHeader = true;

    /* 表头行索引 */
    public int headerRowNumber = 0;

    @Override
    public boolean isCreateHeader() {
        return isCreateHeader;
    }

    @Override
    public int getHeaderRowNumber() {
        return headerRowNumber;
    }

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        //1.0 设置背景色:
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());// 设置背景色
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //2.0 设置边框:
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框

        //3.0 设置居中:
        cellStyle.setAlignment(HorizontalAlignment.CENTER); // 居中

        //4.0 设置字体:
        Font font = workbook.createFont();
        font.setFontName("楷体");
        font.setFontHeightInPoints((short) 10);//设置字体大小
        cellStyle.setFont(font);//选择需要用到的字体格式

        return cellStyle;
    }
}
