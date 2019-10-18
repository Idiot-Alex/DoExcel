package com.hotstrip.DoExcel.excel;

import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ColumnProperty {
    // 字段
    private Field field;
    // 字段值
    private String value;
    // 格式化之后的值
    private String formatValue;
    // 字段 get 方法
    private Method getMethod;
    // 字段所在单元格最大宽度
    private int width;
    // 单元格样式
    private CellStyle cellStyle;
    // 属性替换对象
    private List<HashMap> rewritePropertyList = new ArrayList<HashMap>();

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getFormatValue() {
        return formatValue;
    }

    public void setFormatValue(String formatValue) {
        this.formatValue = formatValue;
    }

    public Method getGetMethod() {
        return getMethod;
    }

    public void setGetMethod(Method getMethod) {
        this.getMethod = getMethod;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public List<HashMap> getRewritePropertyList() {
        return rewritePropertyList;
    }

    public void setRewritePropertyList(List<HashMap> rewritePropertyList) {
        this.rewritePropertyList = rewritePropertyList;
    }
}
