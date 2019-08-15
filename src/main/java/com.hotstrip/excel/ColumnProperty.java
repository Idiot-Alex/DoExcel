package com.hotstrip.excel;

import java.lang.reflect.Field;

public class ColumnProperty {
    // 字段
    private Field field;
    // 名称
    private String name;

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
