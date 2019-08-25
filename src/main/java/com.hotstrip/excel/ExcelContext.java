package com.hotstrip.excel;

import com.hotstrip.annotation.DoSheet;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.util.Const;
import com.hotstrip.excel.util.DoExcelUtil;
import com.hotstrip.exception.DoExcelException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
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

    @Deprecated
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

            String cellValue = columnProperty.getName();

            // 是否需要国际化
            if (this.resourceBundle != null) {
                try {
                    cellValue = new String(this.resourceBundle.getString(cellValue).getBytes("ISO-8859-1"), "UTF-8");
                } catch (UnsupportedEncodingException e) {

                } catch (MissingResourceException e) {
                    logger.warn("the key [{}] is not match value", cellValue);
                }
            }

            // 设置单元格宽度 为 0 就设置自动适应表头
            doHeadRowColumnWidth(cellNum, columnProperty, cellValue);

            Cell cell = row.createCell(cellNum);
            cell.setCellValue(cellValue);
        }
    }

    /**
     * 设置表头行 单元格宽度
     * 为 0 就设置自动适应表头
     * 每个单元格都设置的话，会影响性能
     * @param cellNum
     * @param columnProperty
     * @param cellValue
     */
    private void doHeadRowColumnWidth(int cellNum, ColumnProperty columnProperty, String cellValue) {
        int cellValueLength = DoExcelUtil.getValueLength(cellValue);
        int width = columnProperty.getWidth();
        /**
         * 判断值的真正长度
         * 设置单元格宽度需要 n * 256 其中 n 表示字符个数
         */
        if (cellValueLength > width) {
            // 设置字段单元格长度
            columnProperty.setWidth(cellValueLength);
            this.currentSheet.setColumnWidth(cellNum, DoExcelUtil.calcCellWidth(columnProperty.getWidth()));
        } else {
            // 自动适应宽度
            if (width == 0)
                this.currentSheet.autoSizeColumn(cellNum, true);
            else {
                columnProperty.setWidth(width);
                this.currentSheet.setColumnWidth(cellNum, DoExcelUtil.calcCellWidth(columnProperty.getWidth()));
            }
        }
    }

    /**
     * 设置内容数据单元格宽度
     * @param cellNum
     * @param columnProperty
     * @param cellValue
     * @param isEndRow
     */
    private void doContentRowColumnWidth(int cellNum, ColumnProperty columnProperty, String cellValue, boolean isEndRow) {
        // 开始的处理逻辑和表头行一样
        int cellValueLength = DoExcelUtil.getValueLength(cellValue);
        int width = columnProperty.getWidth();
        if (cellValueLength > width) {
            // 设置字段单元格长度
            columnProperty.setWidth(cellValueLength);
        }
        // 如果是最后一行 设置单元格宽度
        if (isEndRow) {
            this.currentSheet.setColumnWidth(cellNum, DoExcelUtil.calcCellWidth(columnProperty.getWidth()));
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
                // 如果单元格样式没有设置
                if (columnProperty.getCellStyle() == null) {
                    columnProperty.setCellStyle(this.workbook.createCellStyle());
                }
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

                // 如果当前属性重写字段值集合大小 > 0
                if (columnProperty.getRewritePropertyList().size() > 0) {
                    cellValue = doRewriteColumnValue(columnProperty, cellValue);
                }

                // 处理单元格宽度
                doContentRowColumnWidth(cellNum, columnProperty, cellValue.toString(), i == list.size() - 1);

                fillCellValue(cell, columnProperty.getCellStyle(), cellValue);
            }
        }
    }

    /**
     * 重写字段值
     * @param columnProperty
     * @param cellValue
     * @return
     */
    private Object doRewriteColumnValue(ColumnProperty columnProperty, Object cellValue) {
        Object tempValue = cellValue;
        // 遍历集合匹配数据
        for (HashMap map : columnProperty.getRewritePropertyList()) {
            // 需要用字符串匹配
            if (map.containsKey(cellValue.toString())) {
                tempValue = map.get(cellValue.toString());
                break;
            }
        }
        return tempValue;
    }

    /**
     * 设置单元格值
     * 判断单元格数据类型
     * 数值型 又分为 整数 小数
     * 文本型 百分比数值 时间格式化 也就是字符串
     * 日期时间型 需要格式化字符串
     * @param cell
     * @param cellStyle
     * @param cellValue
     */
    private void fillCellValue(Cell cell, CellStyle cellStyle, Object cellValue) {
        // 格式化
        DataFormat dataFormat = this.workbook.createDataFormat();

        boolean isNum = DoExcelUtil.isNum(cellValue);
        boolean isInteger = DoExcelUtil.isIntger(cellValue);
        boolean isPercent = DoExcelUtil.isPercent(cellValue);
        // 整数
        if (isNum && isPercent == false) {
            // 整数不显示小数
            if (isInteger)
                cellStyle.setDataFormat(dataFormat.getFormat("0"));
            else
                cellStyle.setDataFormat(dataFormat.getFormat("0.00"));
            // 设置样式
            cell.setCellStyle(cellStyle);
            cell.setCellValue(Double.parseDouble(cellValue.toString()));
        } else {
            // 日期时间格式 JDK 1.8 之后的日期格式暂时没实现
            if (cellValue instanceof Date) {
                cellStyle.setDataFormat(dataFormat.getFormat(Const.YMDHMS));
                cell.setCellStyle(cellStyle);
                cell.setCellValue((Date) cellValue);
            } else {
                // 字符串
                cellStyle.setDataFormat(dataFormat.getFormat("@"));
                cell.setCellStyle(cellStyle);
                cell.setCellValue(cellValue.toString());
            }
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

        // 设置表标题
        String title = doSheet != null ? doSheet.title() : Const.UNTITLED;
        this.currentSheet = DoExcelUtil.createSheet(this.workbook, title);

        // 设置国际化资源
        boolean localeFlag = true;
        if (this.locale == null) {
            logger.warn("the locale is null, cannot config resourceBundle");
            localeFlag = false;
        }
        if (doSheet.localeResource().length() < 1) {
            logger.warn("the localeResource is invalid, please check it, and value is [{}]", doSheet.localeResource());
            localeFlag = false;
        }
        if (localeFlag) {
            ResourceBundle resourceBundle = ResourceBundle.getBundle(doSheet.localeResource(), this.locale);
            this.resourceBundle = resourceBundle;
        }
    }

    /**
     * 添加行数据
     * @param list
     */
    public void addRow(List<Object> list) {
        // 单元格样式
        CellStyle cellStyle = this.workbook.createCellStyle();

        Row row = getNextRow();
        // 遍历数据
        for (int i = 0; i < list.size(); i++) {
            int cellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
            Cell cell = row.createCell(cellNum);

            // 非空判断
            Object cellValue = list.get(i) == null ? "" : list.get(i);

            fillCellValue(cell, cellStyle, cellValue);
        }
    }

    /**
     * 完成写入文件
     * 关闭流
     */
    public void close() {
        try {
            this.workbook.write(this.outputStream);
            this.workbook.close();
            this.outputStream.close();
        } catch (IOException e) {
            throw new DoExcelException("IO Exception for close workbook", e);
        }
    }

}
