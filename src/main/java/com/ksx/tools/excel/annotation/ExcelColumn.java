package com.ksx.tools.excel.annotation;

import com.ksx.tools.excel.format.DataFormat;
import com.ksx.tools.excel.format.DefaultDataFormat;
import com.ksx.tools.excel.style.DataCellStyle;
import com.ksx.tools.excel.style.DefaultDataCellStyle;
import com.ksx.tools.excel.style.DefaultHeaderCellStyle;
import com.ksx.tools.excel.style.HeaderCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 字段属性注解
 * Created by kangshuangxi on 2016/12/29.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, })
public @interface ExcelColumn {

    /**
     * 导出到Excel中的名字
     * @return
     */
    String name() default "";

    /**
     * 配置列的名称,对应A,B,C,D....
     * @return
     */
    String column();

    /**
     * 提示信息
      * @return
     */
    String prompt() default "";

    /**
     * 设置只能选择不能输入的列内容（暂未实现）
     * @return
     */
    String[] combo() default {};

    /**
     * 设置本字数据是否导出，默认为true导出
     * @return
     */
    boolean isExport() default true;

    /**
     * 单元格数据显示样式
     * @return
     */
    String dataPattern() default "";

    /**
     * excel 表头单元格格式
     * @return
     */
    Class<? extends HeaderCellStyle> headerCellStyle() default DefaultHeaderCellStyle.class;

    /**
     * 数据单元格格式
     * @return
     */
    Class<? extends DataCellStyle> dataCellStyle() default DefaultDataCellStyle.class;

    /**
     * 数据格式化
     * @return
     */
    Class<? extends DataFormat> dataFormat() default DefaultDataFormat.class;
}