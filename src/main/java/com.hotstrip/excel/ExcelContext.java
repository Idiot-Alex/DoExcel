package com.hotstrip.excel;

import com.hotstrip.annotation.Column;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.util.DoExcelUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * Excel config
 * such as Locale support
 */
public class ExcelContext {
    // locale
    private Locale locale;
    // resourceBundle
    private ResourceBundle resourceBundle;
    // Workbook title
    private String title;
    private Sheet currentSheet;
    private Workbook workbook;
    // 输出流
    private OutputStream outputStream;

    // 表头
    HeadRow headRow;

    public ExcelContext(ExcelTypeEnums excelTypeEnums) {
        this.workbook = DoExcelUtil.createWorkBook(excelTypeEnums);
    }

    public ExcelContext(OutputStream outputStream, ExcelTypeEnums excelTypeEnums) {
        this(excelTypeEnums);
        this.outputStream = outputStream;
    }


    public Locale getLocale() {
        return locale;
    }

    public void setLocale(Locale locale) {
        this.locale = locale;
    }

    public ResourceBundle getResourceBundle() {
        return resourceBundle;
    }

    public void setResourceBundle(ResourceBundle resourceBundle) {
        this.resourceBundle = resourceBundle;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public HeadRow getHeadRow() {
        return headRow;
    }

    public void setHeadRow(HeadRow headRow) {
        this.headRow = headRow;
    }

    /**
     * 添加数据
     * @param list
     * @param clazz
     */
    public void addContent(List list, Class clazz) {
        this.headRow = new HeadRow(clazz);

        for (int i = 0; i < list.size(); i++) {
//            Row row = DoExcelUtil.createRow()
        }
    }
}
