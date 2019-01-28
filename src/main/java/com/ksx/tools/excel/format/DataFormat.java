package com.ksx.tools.excel.format;

/**
 * 数据格式
 * Created by kangshuangxi on 2017/1/2.
 */
public interface DataFormat<T> {

    /**
     * 数据格式化
     * @param data
     * @return
     */
    String format(T data);
}
