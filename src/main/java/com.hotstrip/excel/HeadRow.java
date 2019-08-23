package com.hotstrip.excel;

import com.hotstrip.annotation.Column;
import com.hotstrip.excel.util.Const;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * 标题行
 */
public class HeadRow {
    // 类
    private Class headClazz;
    // 列表头
    private List<ColumnProperty> columnPropertyList = new ArrayList<ColumnProperty>();

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
                // 设置字段的相关属性
                columnProperty = new ColumnProperty();
                columnProperty.setName(column.name());
                columnProperty.setField(field);
                columnProperty.setGetMethod(getFieldMethod(clazz, field));
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

    /**
     * 获取字段 get 方法
     * 倘若没有这个字段的 get 方法 返回 null
     * @param clazz
     * @param field
     * @return
     */
    private Method getFieldMethod(Class clazz, Field field) {
        String fieldName = field.getName();
        StringBuffer getMethodName = new StringBuffer();
        // 这里需要判断 字段类型 如果是 boolean 用 is 开头
        if (Const.BOOLEAN_LOWER.equals(field.getType().getName())) {
            getMethodName.append(Const.IS);
        } else {
            getMethodName.append(Const.GET);
        }
        getMethodName.append(fieldName.substring(0, 1)
                .toUpperCase())
                .append(fieldName.substring(1));

        try {
            return clazz.getMethod(getMethodName.toString(), new Class[] {});
        } catch (NoSuchMethodException e) {
            return null;
        }
    }
}
