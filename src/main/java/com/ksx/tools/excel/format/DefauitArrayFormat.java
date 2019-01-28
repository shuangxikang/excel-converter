package com.ksx.tools.excel.format;

import java.util.Collection;

/**
 * Created by kangshuangxi on 2017/1/2.
 */
public class DefauitArrayFormat<T extends Collection> implements DataFormat<T> {

    @Override
    public String format(T data) {
        return null;
    }
}
