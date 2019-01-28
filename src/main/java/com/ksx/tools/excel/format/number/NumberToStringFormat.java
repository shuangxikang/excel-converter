package com.ksx.tools.excel.format.number;

import java.text.DecimalFormat;

/**
 * Created by kangshuangxi on 2017/10/13.
 */
public class NumberToStringFormat<T extends Number> implements NumberFormat<T> {

    /* 数字格式为字符串 */
    private static DecimalFormat decimalFormat;

    @Override
    public DecimalFormat getDecimalFormat() {
        if (decimalFormat == null)
            decimalFormat = new DecimalFormat(getPattern());

        return decimalFormat;
    }
}
