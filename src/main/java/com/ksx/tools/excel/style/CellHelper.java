package com.ksx.tools.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Excel cell 样式
 * Created by kangshuangxi on 2016/12/30.
 */
public interface CellHelper {

    /**
     * 获取单元格样式
     * @param workbook  工作簿
     * @return
     */
    CellStyle getCellStyle(Workbook workbook);

}
