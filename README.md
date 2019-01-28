# excel-converter
## 项目说明
> * 本项目是基于POI进行简单封装
> * 意在能够通过注解方式来实现Excel文件和Java Bean之间的相互转换
> * 对Excel文件操作不需要直接关注POI底层API细节
> * 对样式、显示格式、数据输出格式化等易变化部分支持自定义

## API使用说明
* 数据对象定义
```
// 参考
导出对象：ExcelExportVO
导入对象：ExcelImportVO
```

* Excel文件导出
``` 
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
```

* Excel文件导入
```
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
```

## Spring MVC导入导出Excel处理实例
* MVC导出
```
@Controller
@RequestMapping("/excel")
public class ExcelConverterController {
    @RequestMapping("/export")
    public ResultVO export(
            queryVO queryVO,
            @RequestParam(required = false, defaultValue = ".xls")String filenameExtension,
            HttpServletResponse response) {
        //1.0 道出表格类型
        ExcelType excelType = ExcelType.getExcelTypeByFilenameExtension(filenameExtension);
        if (excelType == null)
            return null; //类型不识别参考ExcelType这里根据具体数据交互类型（json、xml、其他）输出错误提示
    
        //1.0 设置HTTP响应头
        String fileName = "ExcelConverterSpringMVCExport" + excelType.getFilenameExtension();
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName);
        response.setContentType(excelType.getContentType());

        //2.0 获取导出数据列表（这里支持多个sheet导出，为使Demo简化这里用单sheet）
        List<ExcelExportVO> excelExportVOList = getBeans(queryVO);

        //3.0 初始化导出信息
        ExcelExport excelExport = new ExcelExport(ExcelExportVO.class);

        //4.0 输出数据
        try {
            excelExport.exportExcel(excelExportVOList, "ExcelConverter导出测试", response.getOutputStream(), new DefaultExcelStyle(excelType));
        } catch (IOException e) {
            LOG.error("数据导出异常：", e);
            return null;    //这里返回适当错误提醒
        }

        return null;
    }
}
```

* MVC导入
```
@Controller
@RequestMapping("/excel")
public class ExcelConverterController {
    @RequestMapping("/imports")
    public ResultVO imports(HttpServletRequest request, 
                @RequestParam(required = false)MultipartFile file,
                @RequestParam(required = false)String sheetName) {
        //1.0 参数检查
        ExcelType excelType = ExcelType.getExcelTypeByFilename(file.getOriginalFilename());
        if (excelType == null)
            return null;    //这里根据具体数据交互类型（json、xml、其他）输出错误提示

        //2.0 初始化导入信息
        ExcelImport excelImport = new ExcelImport(ExcelImportVO.class);
        
        //3.0 读取并解析Excel文件数据
        List<ExcelImportVO> excelImportVOS;
        try {
            //3.1 这里sheetName为空默认第一个sheet
            excelImportVOS = excelImport.importExcel(file.getInputStream(), excelType, sheetName);
        } catch (IOException e) {
            log.error("读取表格异常！", e);
            return null;    //这里输出错误提示
        }

        //3.0 这里处理解析结果

        return null;
    }
}
```