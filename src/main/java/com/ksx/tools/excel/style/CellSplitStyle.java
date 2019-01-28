package com.ksx.tools.excel.style;

import java.lang.reflect.Field;
import java.util.Map;

/**
 * Created by kangshuangxi on 2017/4/26.
 */
public interface CellSplitStyle {

    Field getField();

    int getBeginColumn();

    int getEndColumn();

    int getRows();

    void setRows(int rows);

    Map<Integer, Field> getItemFields();
}
