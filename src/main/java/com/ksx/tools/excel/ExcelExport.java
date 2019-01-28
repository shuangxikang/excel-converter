package com.ksx.tools.excel;

import com.ksx.tools.excel.annotation.ExcelColumn;
import com.ksx.tools.excel.format.DataFormat;
import com.ksx.tools.excel.format.DefaultDataFormat;
import com.ksx.tools.excel.style.CellSplitStyle;
import com.ksx.tools.excel.style.ExcelStyle;
import com.ksx.tools.excel.utils.ReflectUtil;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Excel导出
 * Created by kangshuangxi on 2016/12/30.
 */
public class ExcelExport<T> {
    private static final Logger LOG = LoggerFactory.getLogger(ExcelExport.class);

    /* 导出对象类型（默认每个sheet对应一个类型，如果不存在则默认使用最后一个类型） */
    private final Class<T>[] clazz;
    /* 工作簿 */
    private Workbook workbook;

    private final CellUtil cellUtil;
    private final List<SheetEntry> sheetEntries;
    public ExcelExport(Class<T> ... clazz) {
        this.clazz = clazz;
        this.cellUtil = new CellUtil();
        this.sheetEntries = new ArrayList<>();
    }

    /**
     * 创建工作表
     * @param dataClazz
     * @param sheetData
     * @param sheetName
     */
    public void addSheet(Class dataClazz, List<T> sheetData, String sheetName) {
        sheetEntries.add(new SheetEntry(dataClazz, sheetData, sheetName));
    }

    /**
     * 将列表数据输出到Excel表格（单个工作表导出）
     * @param sheetData     导出数据列表
     * @param sheetName     工作表的名称
     * @param output        文件输出流
     * @param excelStyle    导出样式
     * @return  写出文件流
     */
    public boolean exportExcel(List<T> sheetData, String sheetName, OutputStream output, ExcelStyle excelStyle) {
        if (sheetData == null || sheetData.size() == 0 || output == null || excelStyle == null)
            return false;

        Class dataClazz;
        if (clazz == null || clazz.length == 0) {
            dataClazz = sheetData.get(0).getClass();
        } else {
            dataClazz = clazz[0];
        }

        addSheet(dataClazz, sheetData, sheetName);

        return exportExcel(output, excelStyle);
    }

    /**
     * 将列表数据输出到Excel表格（多个工作表导出）
     * @param output
     * @param excelStyle
     * @return
     */
    public boolean exportExcel(OutputStream output, ExcelStyle excelStyle) {
        if (sheetEntries.size() == 0)
            return false;

        //1.0 创建工作薄对象
        workbook = Excel.createWorkbook(excelStyle.getExcelType());

        //2.0 创建工作表
        for (SheetEntry sheetEntry : sheetEntries) {
            Sheet sheet = Excel.createSheet(workbook, sheetEntry.getSheetName());
            createExcelDataSheet(getFieldMap(sheetEntry.getDataClazz()), sheet, sheetEntry.getSheetData(), excelStyle);
        }
        try {
            workbook.write(output);
            output.flush();
            output.close();
            return true;
        } catch (IOException e) {
            LOG.error("表格导出异常！", e);
        }
        return false;
    }


    /**
     * 创建数据表
     * @param fieldsMap
     * @param sheet
     * @param sheetDataList
     * @param excelStyle
     * @return
     */
    protected boolean createExcelDataSheet(Map<Integer, Field> fieldsMap, Sheet sheet, List<T> sheetDataList, ExcelStyle excelStyle) {
        if (fieldsMap == null || fieldsMap.size() == 0 || sheet == null || sheetDataList == null)
            return false;

        //1.0 设置表头
        createExcelHeader(fieldsMap, sheet, excelStyle.getStartHeaderRowNumber());

        //2.0 设置表格数据
        return createExcelDataRows(fieldsMap, sheetDataList, sheet, excelStyle.getStartDataRowNumber());
    }

    /**
     * 写出列头字段
     * @param fieldsMap
     * @param sheet
     * @param startHeaderRowNumber
     * @return
     */
    protected boolean createExcelHeader(Map<Integer, Field> fieldsMap, Sheet sheet, int startHeaderRowNumber) {
        if (fieldsMap == null || fieldsMap.size() == 0 || sheet == null)
            return false;

        Row row = sheet.createRow(startHeaderRowNumber);
        Cell cell;
        Map<Integer, Field> headerFields = ReflectUtil.getAllHeaderFiled(fieldsMap);
        for (Map.Entry entry : headerFields.entrySet()) {
            if (entry == null || entry.getKey() == null || entry.getValue() == null)
                continue;

            ExcelColumn excelColumn = ((Field)entry.getValue()).getAnnotation(ExcelColumn.class);
            cell = row.createCell((int)entry.getKey());

            setHeaderCellStyle((Field)entry.getValue(), cell);
            cell.setCellValue(excelColumn.name());
        }

        return true;
    }

    /**
     * 创建数据表
     * @param fieldsMap             输出字段列表
     * @param sheetDataList         数据数据集
     * @param sheet                 数据表
     * @param startDataRowNumber    数据起始行
     * @return
     */
    private boolean createExcelDataRows(Map<Integer, Field> fieldsMap, List<T> sheetDataList, Sheet sheet, int startDataRowNumber) {
        //1.0 参数检查
        if (fieldsMap == null || fieldsMap.size() == 0 || sheet == null || sheetDataList == null)
            return false;

        CellSplitStyle cellSplitStyle = ReflectUtil.getSplitField(fieldsMap);

        //2.0 写出 sheet 数据
        int dataRowIndex = startDataRowNumber;
        for (int i = 0; i < sheetDataList.size(); i++) {
            Row row = sheet.createRow(dataRowIndex++);
            T data = sheetDataList.get(i);

            if (cellSplitStyle == null) {
                createExcelDataRow(fieldsMap, row, data);
            } else {
                Object splitFieldValue = ReflectUtil.getFiledValue(cellSplitStyle.getField(), data);
                int rows = ReflectUtil.getCollectionSize(splitFieldValue);
                cellSplitStyle.setRows(rows);
                createExcelDataRow(fieldsMap, cellSplitStyle, row, data);

                dataRowIndex += (rows -1);
            }
        }

        //3.0 设置单元格提示
        setCellPrompt(fieldsMap, sheet, startDataRowNumber, sheetDataList.size());
        return true;
    }

    /**
     * 写出数据行，并设置样式及数据
     * @param fieldsMap 输出字段列表
     * @param row       表格行
     * @param date      待输出数据
     * @return
     */
    private boolean createExcelDataRow(Map<Integer, Field> fieldsMap, Row row, Object date) {
        if (fieldsMap == null || fieldsMap.size() == 0 || row == null || date == null)
            return false;

        Cell cell;
        for (Map.Entry entry : fieldsMap.entrySet()) {
            Field field = (Field) entry.getValue();
            cell = row.createCell((int)entry.getKey());

            setCellProperties(field, date, cell);
        }

        return true;
    }

    /**
     * 写出数据行，并设置样式及数据，处理单元格合并
     * @param fieldsMap
     * @param cellSplitStyle
     * @param row
     * @param data
     * @return
     */
    private boolean createExcelDataRow(Map<Integer, Field> fieldsMap, CellSplitStyle cellSplitStyle, Row row, T data) {
        if (fieldsMap == null || fieldsMap.size() == 0 || row == null || data == null)
            return false;

        Object splitFieldValue = ReflectUtil.getFiledValue(cellSplitStyle.getField(), data);
        Cell cell;
        int columnIndex;
        for (Map.Entry entry : fieldsMap.entrySet()) {
            Field field = (Field) entry.getValue();
            columnIndex = (int) entry.getKey();
            cell = row.createCell(columnIndex);

            if (cellSplitStyle.getBeginColumn() != columnIndex) {
                Excel.setCellRangeAddress(row, columnIndex, ReflectUtil.getCollectionSize(splitFieldValue), 0);
                setCellProperties(field, data, cell);
            } else {
                setCellSplit(row, cellSplitStyle, splitFieldValue);
            }
        }

        return true;
    }

    /**
     * 设置拆分单元格格式
     * @param row
     * @param cellSplitStyle
     * @param fieldValue
     */
    private void setCellSplit(Row row, CellSplitStyle cellSplitStyle, Object fieldValue) {
        int beginRowIndex = row.getRowNum();
        Sheet sheet = row.getSheet();

        Row cellSplitRow = row;
        int itemIndex = 0;
        for (Object obj : (Collection)fieldValue) {
            if (itemIndex == 0) {
                itemIndex++;
            } else {
                cellSplitRow = sheet.createRow(beginRowIndex + itemIndex);
                itemIndex++;
            }

            createExcelDataRow(cellSplitStyle.getItemFields(), cellSplitRow, obj);
        }
    }

    /**
     * 设置单元格属性
     * @param field
     * @param data
     * @param cell
     */
    private void setCellProperties(Field field, Object data, Cell cell) {
        if (field == null || data == null || cell == null)
            return;

        //1.0 设置单元格样式
        setDataCellStyle(field, cell);
        setCellDataFormat(field, cell);

        //2.0 设置单元格内容
        ExcelColumn column = field.getAnnotation(ExcelColumn.class);
        if (column.isExport())
            setCellData(field, data, cell);
    }

    /**
     * 设置单元格值
     * @param field 字段
     * @param data  数据对象
     * @param cell  单元格
     */
    private void setCellData(Field field, Object data, Cell cell) {
        try {
            //1.0 获取字段值
            Object fieldValue = field.get(data);
            if (fieldValue == null)
                return;

            //2.0 是否需要特殊格式化处理
            if (isNotDefaultDataFormat(field)) {
                cell.setCellValue(formatData(field, fieldValue));
                return;
            }

            //3.0 根据字段值类型设置单元格 value，属性的复杂类型暂不处理直接调用toString
            if (fieldValue instanceof Number) {
                cell.setCellValue(((Number) fieldValue).doubleValue());
            } else if (fieldValue instanceof Date) {
                cell.setCellValue((Date) fieldValue);
            } else if (fieldValue instanceof Boolean) {
                cell.setCellValue((Boolean) fieldValue);
            } else if (fieldValue instanceof String) {
                cell.setCellValue((String) fieldValue);
            } else{
                cell.setCellValue(formatData(field, fieldValue));
            }
        } catch (Exception e) {
            LOG.error("写入单元格数据异常！", e);
        }
    }

    /**
     * 数据格式化
     * @param field
     * @param date
     * @param <D>
     * @return
     */
    private <D> String formatData(Field field, D date) {
        DataFormat<D> dataFormat = cellUtil.getDataFormat(field);
        try {
            return dataFormat.format(date);
        } catch (Exception e) {
            LOG.error("单元格数据格式化异常！", e);
        }

        return null;
    }

    /**
     * 数据格式化类型是否为默认格式化
     * 如果为默认格式化的基本类型或者基本类型封装类包含 String 都按普通方式处理
     * 如果指定了自定义数据格式化类型则直接调用格式化方法进行参数格式化然后设置
     * @param field 属性字段
     * @return
     */
    private boolean isNotDefaultDataFormat(Field field) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if (excelColumn.dataFormat().isAssignableFrom(DefaultDataFormat.class))
            return false;

        return true;
    }

    /**
     * 根据注解参数设置数据单元格样式
     * @param field
     * @param cell
     */
    private void setDataCellStyle(Field field, Cell cell) {
        try {
            CellStyle cellStyle = cellUtil.getDataCellStyle(field, cell.getSheet().getWorkbook());
            cell.setCellStyle(cellStyle);
        } catch (Exception e) {
            LOG.error("ExcelExport.setDataCellStyle error！", e);
        }
    }

    /**
     * 根据注解参数设置表头单元格样式
     * @param field
     * @param cell
     */
    private void setHeaderCellStyle(Field field, Cell cell) {
        try {
            CellStyle cellStyle = cellUtil.getHeaderCellStyle(field, cell.getSheet().getWorkbook());
            cell.setCellStyle(cellStyle);
        } catch (Exception e) {
            LOG.error("ExcelExport.setHeaderCellStyle error！", e);
        }
    }

    /**
     * 设置单元格数据展示样式
     * @param field
     * @param cell
     */
    private void setCellDataFormat(Field field, Cell cell) {
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if ("".equals(excelColumn.dataPattern()))
            return;

        CellStyle style = cell.getCellStyle();
        if (style == null)
            style = workbook.createCellStyle();

        style.setDataFormat(workbook.createDataFormat().getFormat(excelColumn.dataPattern()));
        cell.setCellStyle(style);
    }

    /**
     * 设置单元格提示
     * @param fieldsMap
     * @param sheet
     * @param startRowNumber
     * @param endRowNumber
     */
    private void setCellPrompt(Map<Integer, Field> fieldsMap, Sheet sheet, int startRowNumber, int endRowNumber) {
        for (Map.Entry entry : fieldsMap.entrySet()) {
            Field field = (Field)entry.getValue();
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            if (column.prompt().equals(""))
                continue;

            int columnStart = (int) entry.getKey(), columnEnd = columnStart;
            Excel.setPrompt(sheet, column.name(), column.prompt(), startRowNumber, endRowNumber, columnStart, columnEnd);
        }
    }

    /**
     * 获取字段列表
     * @param clazz
     * @return
     */
    private Map<Integer, Field> getFieldMap(Class clazz) {
        return ReflectUtil.getFieldMap(clazz);
    }

    /**
     * 单元格格式
     */
    private class CellUtil {
        /* 单元格样式 */
        private static final String DATA_CELL_STYLE_PREFIX = "data_style_";
        private static final String HEADER_CELL_STYLE_PREFIX = "header_style_";
        private ThreadLocal<ConcurrentHashMap<String, CellStyle>> threadLocalStyle = new ThreadLocal();

        /* 单元格数据格式化 */
        private ThreadLocal<ConcurrentHashMap<String, DataFormat>> threadLocalDataFormat = new ThreadLocal();

        /**
         * 数据单元格样式
         * @param field
         * @param workbook
         * @return
         */
        private CellStyle getDataCellStyle(Field field, Workbook workbook) throws IllegalAccessException, InstantiationException {
            String dataCellStyleKey = DATA_CELL_STYLE_PREFIX + field.getDeclaringClass().getSimpleName() + field.getName();
            CellStyle cellStyle = getCellStyle(dataCellStyleKey);

            if (cellStyle == null) {
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                cellStyle = column.dataCellStyle().newInstance().getCellStyle(workbook);
                putCellStyle(dataCellStyleKey, cellStyle);
            }

            return cellStyle;
        }

        /**
         * 获取标题单元格样式
         * @param field
         * @param workbook
         * @return
         */
        private CellStyle getHeaderCellStyle(Field field, Workbook workbook) throws IllegalAccessException, InstantiationException {
            String headerCellStyleKey = HEADER_CELL_STYLE_PREFIX + field.getDeclaringClass().getSimpleName() + field.getName();
            CellStyle cellStyle = getCellStyle(headerCellStyleKey);

            if (cellStyle == null) {
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                cellStyle = column.headerCellStyle().newInstance().getCellStyle(workbook);
                putCellStyle(headerCellStyleKey, cellStyle);
            }

            return cellStyle;
        }

        /**
         * 根据 key 获取单元格样式
         * @param key
         * @return
         */
        private CellStyle getCellStyle(String key) {
            ConcurrentHashMap<String, CellStyle> cellStyleMap = threadLocalStyle.get();
            if (cellStyleMap == null) {
                cellStyleMap = new ConcurrentHashMap<>();
                threadLocalStyle.set(cellStyleMap);
            }

            return cellStyleMap.get(key);
        }

        /**
         * 缓存 cellStyle
         * @param key
         * @param cellStyle
         */
        private void putCellStyle(String key, CellStyle cellStyle) {
            threadLocalStyle.get().put(key, cellStyle);
        }

        /**
         * 单元格数据格式
         * @param field
         * @return
         */
        public DataFormat getDataFormat(Field field) {
            String filedMark = field.toGenericString();
            ConcurrentHashMap<String, DataFormat> dataFormats = threadLocalDataFormat.get();
            if (dataFormats == null) {
                dataFormats = new ConcurrentHashMap<>();
                threadLocalDataFormat.set(dataFormats);
            }

            DataFormat dataFormat = dataFormats.get(filedMark);
            if (dataFormat == null) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                try {
                    dataFormat = excelColumn.dataFormat().newInstance();
                    dataFormats.put(filedMark, dataFormat);
                } catch (Exception e) {
                    LOG.error("单元格数据格式化类型");
                }
            }

            return dataFormat;
        }
    }

    /**
     * 工作表
     * @param <T>
     */
    private class SheetEntry<T> {
        /* 数据类型 */
        private Class dataClazz;
        /* 工作表数据列表 */
        private List<T> sheetData;
        /* 工作表名称 */
        private String sheetName;
        SheetEntry(Class dataClazz, List<T> sheetData, String sheetName) {
            this.dataClazz = dataClazz;
            this.sheetData = sheetData;
            this.sheetName = sheetName;
        }

        public Class getDataClazz() {
            return dataClazz;
        }

        public void setDataClazz(Class dataClazz) {
            this.dataClazz = dataClazz;
        }

        public List<T> getSheetData() {
            return sheetData;
        }

        public void setSheetData(List<T> sheetData) {
            this.sheetData = sheetData;
        }

        public String getSheetName() {
            return sheetName;
        }

        public void setSheetName(String sheetName) {
            this.sheetName = sheetName;
        }
    }
}