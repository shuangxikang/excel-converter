package com.ksx.tools.excel.format.number;

import com.ksx.tools.excel.format.DataFormat;

import java.text.DecimalFormat;

/**
 *
 * Created by kangshuangxi on 2017/4/8.
 */
public interface NumberFormat<T extends Number> extends DataFormat<T> {

    @Override
    default String format(T data) {
        if (data != null) {
            DecimalFormat decimalFormat = getDecimalFormat();
            return decimalFormat.format(data);
        }

        return null;
    }

    /**
     * 模式字符串
     * @return
     */
    default String getPattern() {
        return "#";
    }


    DecimalFormat getDecimalFormat();
}
