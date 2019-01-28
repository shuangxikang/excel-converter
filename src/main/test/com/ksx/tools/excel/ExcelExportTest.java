package com.ksx.tools.excel;

import com.ksx.tools.excel.style.DefaultExcelStyle;
import com.ksx.tools.excel.style.ExcelStyle;
import com.ksx.tools.excel.utils.ExcelType;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.sql.Date;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by kangshuangxi on 2017/1/17.
 */
public class ExcelExportTest {

    private List<ExcelExportVO> getBeans() {
        List<ExcelExportVO> excelExportVOList = new ArrayList();

        ExcelExportVO excelExportVO = new ExcelExportVO();
        excelExportVO.setUserAccount("张三");
        excelExportVO.setAge(29);
        excelExportVO.setGender('M');
        excelExportVO.setBirthday(Date.valueOf("1990-10-10"));
        excelExportVO.setPassword("12345");
        excelExportVO.setAmount(new BigDecimal(11000.12));
        excelExportVO.setIdCard(123123123123123123L);
        List<ExcelExportVO.Address> addressListZF = new ArrayList<>();
        addressListZF.add(new ExcelExportVO.Address("张先生", "13088888888", "浙江省杭州市西湖区XXX"));
        addressListZF.add(new ExcelExportVO.Address("李先生", "13066666666", "浙江省杭州市余杭区XXX"));
        excelExportVO.setAddressList(addressListZF);
        excelExportVO.setStatus(2);
        excelExportVOList.add(excelExportVO);

        ExcelExportVO excelExportVO2 = new ExcelExportVO();
        excelExportVO2.setUserAccount("李四");
        excelExportVO2.setAge(26);
        excelExportVO2.setGender('F');
        excelExportVO2.setBirthday(Date.valueOf("1994-10-10"));
        excelExportVO2.setPassword("54321");
        excelExportVO2.setAmount(new BigDecimal(5100.35));
        excelExportVO2.setIdCard(321321321321321321L);
        List<ExcelExportVO.Address> addressListLS = new ArrayList<>();
        addressListLS.add(new ExcelExportVO.Address("王先生", "13077777777", "河南省南阳市卧龙区XXX"));
        excelExportVO2.setAddressList(addressListLS);
        excelExportVO2.setStatus(1);
        excelExportVOList.add(excelExportVO2);

        return excelExportVOList;
    }

    @Test
    public void testExport() throws FileNotFoundException {
        //1.0 文件输出流（这里我直接保存到本地当前用户工作目录）
        File file = new File(System.getProperty("user.dir") + "/ExcelConverterExport.xls");
        FileOutputStream fos = new FileOutputStream(file);

        //2.0 获取导出数据列表
        List<ExcelExportVO> excelExportVOList = getBeans();

        //3.0 构造导出对象
        ExcelExport excelExport = new ExcelExport();

        //4.0 构建导出样式（一般不需定制导出样式等特殊样式，直接使用DefaultExcelStyle即可，ExcelType知道导出 Excel 类型.xls或.xlsx）
        ExcelStyle excelStyle = new DefaultExcelStyle(ExcelType.xls);

        //5.0 导出 Excel数据
        excelExport.exportExcel(excelExportVOList,"用户信息", fos, excelStyle);
    }
}
