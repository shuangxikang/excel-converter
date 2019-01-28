package com.ksx.tools.excel.format;

/**
 * 默认数据格式化
 * Created by kangshuangxi on 2017/1/2.
 */
public class DefaultDataFormat<T> implements DataFormat<T> {

    @Override
    public String format(T data) {
        return data.toString();
    }

}
