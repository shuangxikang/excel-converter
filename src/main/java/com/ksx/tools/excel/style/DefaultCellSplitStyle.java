package com.ksx.tools.excel.style;

import java.lang.reflect.Field;
import java.util.Map;

/**
 * Created by kangshuangxi on 2017/4/26.
 */
public class DefaultCellSplitStyle implements CellSplitStyle {

    private Field field;

    private final int beginColumn;

    private final int endColumn;

    private int rows;

    private final Map<Integer, Field> itemFields;

    public DefaultCellSplitStyle(Field field, Map<Integer, Field> itemFields, int beginColumn, int endColumn) {
        this.field = field;
        this.itemFields = itemFields;
        this.beginColumn = beginColumn;
        this.endColumn = endColumn;
    }

    @Override
    public Field getField() {
        return field;
    }

    @Override
    public int getBeginColumn() {
        return beginColumn;
    }

    @Override
    public int getEndColumn() {
        return endColumn;
    }

    @Override
    public int getRows() {
        return rows;
    }

    @Override
    public Map<Integer, Field> getItemFields() {
        return itemFields;
    }

    @Override
    public void setRows(int rows) {
        this.rows = rows;
    }
}
