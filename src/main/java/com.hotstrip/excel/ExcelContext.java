package com.hotstrip.excel;

import com.hotstrip.annotation.DoSheet;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.util.DoExcelUtil;
import com.hotstrip.exception.DoExcelException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Excel config
 * such as Locale support
 */
public class ExcelContext {
    private static Logger logger = LoggerFactory.getLogger(ExcelContext.class);
    // locale
    private Locale locale;
    // resourceBundle
    private ResourceBundle resourceBundle;
    // 当前 Sheet
    private Sheet currentSheet;
    // 当前行号
    private int currentRowNum;
    // 工作薄
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

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public int getCurrentRowNum() {
        return currentRowNum;
    }

    public void setCurrentRowNum(int currentRowNum) {
        this.currentRowNum = currentRowNum;
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
        // 添加 Sheet
        addSheet(clazz);

        // 添加标题行
        addHeadRow(clazz);

        // 写入内容
        for (int i = 0; i < list.size(); i++) {
            Row row = getNextRow();
            int cellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
            Cell cell = row.createCell(cellNum);
            cell.setCellValue(list.get(i).toString());
        }
    }

    /**
     * 添加标题行
     * @param clazz
     */
    private void addHeadRow(Class clazz){
        this.headRow = new HeadRow(clazz);

        Row row = getNextRow();
        for (ColumnProperty columnProperty : this.headRow.getColumnPropertyList()) {
            int cellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
            Cell cell = row.createCell(cellNum);
            cell.setCellValue(columnProperty.getName());
        }
    }

    /**
     * 获取下一行数据
     * 当前行号自增
     * @return
     */
    private Row getNextRow() {
        Row row = DoExcelUtil.createRow(this.currentSheet, this.currentRowNum);
        this.currentRowNum++;
        return row;
    }

    /**
     * 添加 Sheet
     * @param clazz
     */
    private void addSheet(Class clazz) {
        // 获取类上的注解
        DoSheet doSheet = (DoSheet) clazz.getAnnotation(DoSheet.class);

        String title = doSheet != null ? doSheet.title() : "Untitled";

        this.currentSheet = DoExcelUtil.createSheet(this.workbook, title);
    }

    /**
     * 完成写入文件
     * 关闭流
     */
    public void close() {
        try {
            this.workbook.write(this.outputStream);
            this.workbook.close();
        } catch (IOException e) {
            throw new DoExcelException("IO Exception for close workbook", e);
        }
    }
}
