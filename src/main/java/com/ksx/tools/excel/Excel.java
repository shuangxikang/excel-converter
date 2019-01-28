package com.ksx.tools.excel;

import com.ksx.tools.excel.utils.ExcelType;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;

/**
 * Excel 工作表
 * Created by kangshuangxi on 2016/12/30.
 */
public class Excel {

    /**
     * 根据类型创建工作簿
     * @param type
     * @return
     */
    public static Workbook createWorkbook(ExcelType type) {
        Workbook workbook = null;
        switch (type) {
            case xls:
                workbook = new HSSFWorkbook();
                break;
            case xlsx:
                workbook = new XSSFWorkbook();
        }

        return workbook;
    }

    /**
     * 根据类型创建工作簿
     * @param type
     * @return
     */
    public static Workbook createWorkbook(InputStream inputStream, ExcelType type) throws IOException {
        Workbook workbook = null;
        switch (type) {
            case xls:
                workbook = new HSSFWorkbook(inputStream);
                break;
            case xlsx:
                workbook = new XSSFWorkbook(inputStream);
        }

        return workbook;
    }

    /**
     * 创建工作表
     * @param workbook
     * @param sheetName
     * @return
     */
    public static Sheet createSheet(Workbook workbook, String sheetName) {
        if (workbook == null)
            return null;

        return workbook.createSheet(sheetName);
    }

    /**
     * 根据工作表类型创建 Cell 样式
     * @param workbook
     * @return
     */
    public static CellStyle createCellStyle(Workbook workbook) {
        if (workbook == null)
            return null;

        CellStyle cellStyle = workbook.createCellStyle();
        return cellStyle;
    }

    /**
     * 将EXCEL中列的字母序A,B,C,D…… 转换成0,1,2,3……表示的数字序
     * @param column    字母序：A、B、C……
     */
    public static int getExcelColumnIndex(String column) {
        //1.0 字符转换为大写表示
        String columnUpper = column.toUpperCase();

        //2.0 声明字母系开始字符内码
        int beginChar = 'A';

        //3.0 按照字符内码转换字母序到数字序
        char[] columnChars = columnUpper.toCharArray();
        int columnIndex = 0;
        for (int i = columnChars.length - 1, j = 1; i >= 0; i--, j *= 26){
            char c = columnChars[i];
            if (c < 'A' || c > 'Z')
                return 0;

            columnIndex += (c - beginChar + 1) * j;
        }

        return  columnIndex - 1;
    }

    /**
     * 设置单元格上提示
     * @param sheet             要设置的sheet.
     * @param promptTitle       标题
     * @param promptContent     内容
     * @param startRowNumber    开始行
     * @param endRowNumber      结束行
     * @param columnStart       开始列
     * @param columnEnd         结束列
     * @return 设置好的sheet.
     */
    public static Sheet setPrompt(Sheet sheet, String promptTitle, String promptContent, int startRowNumber, int endRowNumber, int columnStart, int columnEnd) {
        Objects.requireNonNull(sheet, "表格对象sheet 不能为空");

        //1.0 构造constraint对象（这里DD1暂时不知道为什么这么写）
        DVConstraint dvConstraint = DVConstraint.createCustomFormulaConstraint("DD1");
        //2.0 设置：起始行、终止行、起始列、终止列
        CellRangeAddressList cellRangeAddressList = new CellRangeAddressList(startRowNumber, endRowNumber, columnStart, columnEnd);
        // 数据有效性对象
        DataValidation dataValidation = getDataValidation(sheet, dvConstraint, cellRangeAddressList);
        dataValidation.createPromptBox(promptTitle, promptContent);
        sheet.addValidationData(dataValidation);
        return sheet;
    }

    /**
     * 根据 Sheet 类型创建校验器
     * @param sheet
     * @param dvConstraint
     * @param cellRangeAddressList
     * @return
     */
    private static DataValidation getDataValidation(Sheet sheet, DVConstraint dvConstraint, CellRangeAddressList cellRangeAddressList) {
        Objects.requireNonNull(sheet, "表格对象sheet 不能为空");

        if (sheet instanceof HSSFSheet) {
            return new HSSFDataValidation(cellRangeAddressList, dvConstraint);
        } else {
            XSSFDataValidationConstraint xssfDataValidationConstraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.FORMULA, dvConstraint.getFormula1());

            XSSFDataValidationHelper validationHelper = new XSSFDataValidationHelper((XSSFSheet)sheet);
            return validationHelper.createValidation(xssfDataValidationConstraint, cellRangeAddressList);
        }
    }

    /**
     * 设置单元格合并
     * @param row           数据行
     * @param columnIndex   列索引
     * @param rows          合并行数
     * @param columns       合并列数
     */
    public static void setCellRangeAddress(Row row, int columnIndex, int rows, int columns) {
        if (rows <= 1)
            return;

        int beginRowIndex = row.getRowNum();
        Sheet sheet = row.getSheet();
        CellRangeAddress cellRangeAddress =new CellRangeAddress(beginRowIndex, (beginRowIndex + rows -1), columnIndex, (columnIndex + columns));
        sheet.addMergedRegion(cellRangeAddress);
    }
}
