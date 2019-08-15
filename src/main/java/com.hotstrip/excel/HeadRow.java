package com.hotstrip.excel;

import com.hotstrip.annotation.Column;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class HeadRow {
    // 类
    private Class headClazz;
    // 列表头
    private List<ColumnProperty> columnPropertyList;

    public HeadRow() {
    }

    public HeadRow(Class clazz) {
        // 获取实体类字段信息
        Class tempClass = clazz;
        List<Field> fieldList = new ArrayList<Field>();
        while (tempClass != null) {
            fieldList.addAll(Arrays.asList(tempClass.getDeclaredFields()));
            // 检查父类
            tempClass = tempClass.getSuperclass();
        }
        // 遍历字段信息 匹配字段注解信息
        for (Field field : fieldList) {
            Column column = field.getAnnotation(Column.class);
            ColumnProperty columnProperty = null;
            if (column != null) {
                columnProperty = new ColumnProperty();
                columnProperty.setName(column.name());
                columnProperty.setField(field);
            }
            // 过滤掉没有注解的字段
            if (columnProperty != null) {
                this.columnPropertyList.add(columnProperty);
            }
        }
    }

    public Class getHeadClazz() {
        return headClazz;
    }

    public void setHeadClazz(Class headClazz) {
        this.headClazz = headClazz;
    }

    public List<ColumnProperty> getColumnPropertyList() {
        return columnPropertyList;
    }

    public void setColumnPropertyList(List<ColumnProperty> columnPropertyList) {
        this.columnPropertyList = columnPropertyList;
    }
}
