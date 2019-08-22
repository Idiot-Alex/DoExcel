package com.hotstrip.excel;

import com.hotstrip.annotation.DoSheet;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.util.Const;
import com.hotstrip.excel.util.DoExcelUtil;
import com.hotstrip.exception.DoExcelException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;

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
    // excel 类型
    private ExcelTypeEnums excelTypeEnums;

    // 表头
    HeadRow headRow;

    public ExcelContext(ExcelTypeEnums excelTypeEnums) {
        this.excelTypeEnums = excelTypeEnums;
        this.workbook = DoExcelUtil.createWorkBook(this.excelTypeEnums);
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

    public ExcelTypeEnums getExcelTypeEnums() {
        return excelTypeEnums;
    }

    public void setExcelTypeEnums(ExcelTypeEnums excelTypeEnums) {
        this.excelTypeEnums = excelTypeEnums;
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

        // 写入内容行
        addContentRows(list);
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
            // 设置单元格宽度 为 0 就设置自动适应表头
            doColumnWidth(cellNum, columnProperty.getWidth());

            Cell cell = row.createCell(cellNum);
            cell.setCellValue(columnProperty.getName());
        }
    }

    /**
     * 设置单元格宽度
     * 为 0 就设置自动适应表头
     * 每个单元格都设置的话，会影响性能
     * @param cellNum
     * @param width
     */
    private void doColumnWidth(int cellNum, Integer width) {
        if (width == 0) {
            this.currentSheet.autoSizeColumn(cellNum, true);
        } else {
            this.currentSheet.setColumnWidth(cellNum, width);
        }
    }

    /**
     * 写入内容行数据
     * @param list
     */
    private void addContentRows(List list) {
        for (int i = 0; i < list.size(); i++) {
            Row row = getNextRow();
            for (ColumnProperty columnProperty : this.headRow.getColumnPropertyList()) {
                Object cellValue = new String();
                Method getMethod = columnProperty.getGetMethod();
                if (getMethod != null) {
                    try {
                        cellValue = getMethod.invoke(list.get(i));
                    } catch (IllegalAccessException e) {

                    } catch (InvocationTargetException e) {

                    }
                }
                int cellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
                Cell cell = row.createCell(cellNum);

                // 非空判断
                cellValue = cellValue == null ? "" : cellValue;

                doColumnType(cell, cellValue);
            }
        }
    }

    /**
     * 判断单元格数据类型
     * 数值型 又分为 整数 小数
     * 文本型 百分比数值 时间格式化 也就是字符串
     * @param cell
     * @param cellValue
     */
    private void doColumnType(Cell cell, Object cellValue) {
        if (cellValue instanceof Date) {
            cellValue = DateFormat.getDateTimeInstance().format(cellValue);
        }
        CellStyle cellStyle = this.workbook.createCellStyle();

        boolean isNum = DoExcelUtil.isNum(cellValue);
        boolean isInteger = DoExcelUtil.isIntger(cellValue);
        boolean isPercent = DoExcelUtil.isPercent(cellValue);
        // 整数
        if (isNum && isPercent == false) {
            DataFormat dataFormat = this.workbook.createDataFormat();
            // 整数不显示小数
            if (isInteger)
                cellStyle.setDataFormat(dataFormat.getFormat("#,#0"));
            else
                cellStyle.setDataFormat(dataFormat.getFormat("#,##0.00"));
            // 设置样式
            cell.setCellStyle(cellStyle);
            cell.setCellValue(Double.parseDouble(cellValue.toString()));
        } else {
            cell.setCellValue(cellValue.toString());
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

        String title = doSheet != null ? doSheet.title() : Const.UNTITLED;

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
