package com.ksx.tools.excel;

import com.ksx.tools.excel.utils.ExcelType;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

/**
 * Created by kangshuangxi on 2017/5/3.
 */
public class ExcelImportTest {

    @Test
    public void importTest() throws FileNotFoundException {
        //1.0 初始化导入信息
        ExcelImport excelImport = new ExcelImport(ExcelImportVO.class);

        //2.0 获取导入文件流
        File file = new File(System.getProperty("user.dir") + "/ExcelConverterExport.xls");
        FileInputStream inputStream = new FileInputStream(file);

        //3.0 导入并解析导入数据
        List<ExcelImportVO>  excelImportVOS = excelImport.importExcel(inputStream, ExcelType.xls);

        //4.0 处理导入数据，这里直接在控制台输出
        for (ExcelImportVO excelImportVO : excelImportVOS) {
            System.out.println(
                    excelImportVO.getUserAccount() + "\t" +
                            excelImportVO.getAge() + "\t" +
                            excelImportVO.getBirthday() + "\t" +
                            excelImportVO.getAmount());
        }
    }
}
