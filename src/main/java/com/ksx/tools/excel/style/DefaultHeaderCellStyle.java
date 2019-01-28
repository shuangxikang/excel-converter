package com.ksx.tools.excel.style;

import org.apache.poi.ss.usermodel.*;

/**
 * Created by kangshuangxi on 2016/12/30.
 */
public class DefaultHeaderCellStyle implements HeaderCellStyle {

    /* 是否创建表头 */
    public boolean isCreateHeader = true;

    /* 表头行索引 */
    public int headerRowNumber = 0;

    @Override
    public CellStyle getCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
//        //1.0 设置背景色:
//        cellStyle.setFillForegroundColor((short) 13);// 设置背景色
//        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

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
//        font.setBold(true);//粗体显示
        cellStyle.setFont(font);//选择需要用到的字体格式

        //5.0 设置列宽:第一个参数代表列id(从0开始),第2个参数代表宽度值  参考 ："2016-12-30"的宽度为2500
        //sheet.setColumnWidth(0, 2500);

        //设置自动换行:
        //cellStyle.setWrapText(true);//设置自动换行
        return cellStyle;
    }

    @Override
    public boolean isCreateHeader() {
        return isCreateHeader;
    }

    @Override
    public int getHeaderRowNumber() {
        return headerRowNumber;
    }
}
