package com.ksx.tools.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel 单元格拆分
 * Created by kangshuangxi on 2017/4/26.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, })
public @interface ExcelColumnSplitCell {
    /**
     * 开始列的名称,对应A,B,C,D....
     * @return
     */
    String beginColumn();

    /**
     * 结束列的名称,对应A,B,C,D....
     * @return
     */
    String endColumn();

    /**
     * 单元格拆分项属性类型
     * @return
     */
    Class itemClass();
}
