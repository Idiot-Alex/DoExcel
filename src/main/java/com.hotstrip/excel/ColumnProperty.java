package com.hotstrip.excel;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class ColumnProperty {
    // 字段
    private Field field;
    // 名称
    private String name;
    // 字段 get 方法
    private Method getMethod;
    // 单元格宽度
    private Integer width;

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

    public Method getGetMethod() {
        return getMethod;
    }

    public void setGetMethod(Method getMethod) {
        this.getMethod = getMethod;
    }

    public Integer getWidth() {
        return width;
    }

    public void setWidth(Integer width) {
        this.width = width;
    }
}
