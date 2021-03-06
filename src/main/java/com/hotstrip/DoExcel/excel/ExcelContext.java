package com.hotstrip.DoExcel.excel;

import com.hotstrip.DoExcel.annotation.DoSheet;
import com.hotstrip.DoExcel.enums.ExcelTypeEnums;
import com.hotstrip.DoExcel.excel.util.Const;
import com.hotstrip.DoExcel.excel.util.DoExcelUtil;
import com.hotstrip.DoExcel.excel.util.StringUtil;
import com.hotstrip.DoExcel.exception.DoExcelException;
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

    /**
     * 添加数据
     * @param list java.util.List
     * @param clazz Class
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

            // 获取国际化之后的值
            String cellValue = handleResourceBundle(columnProperty.getValue());
            // 移除首尾空格
            cellValue = StringUtil.lTrim(StringUtil.rTrim(cellValue));

            // 设置单元格宽度 为 0 就设置自动适应表头
            doHeadRowColumnWidth(cellNum, columnProperty, cellValue);

            Cell cell = row.createCell(cellNum);
            cell.setCellValue(cellValue);
        }
    }

    /**
     * 处理国际化
     * @param value
     * @return
     */
    private String handleResourceBundle(String value) {
        // 是否需要国际化
        if (this.resourceBundle != null) {
            try {
                return new String(this.resourceBundle.getString(value).getBytes("ISO-8859-1"), "UTF-8");
            } catch (UnsupportedEncodingException e) {

            } catch (MissingResourceException e) {
                if (logger.isDebugEnabled())
                    logger.debug("the key [{}] is not match value", value);
            }
        }
        return value;
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
        // 计算 列字段值长度
        int cellValueLength = StringUtil.getValueLength(cellValue);
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
        // 若格式化字符串为空就计算真实值的长度
        if (columnProperty.getFormatValue() == null || columnProperty.getFormatValue().length() == 0){
            int cellValueLength = StringUtil.getValueLength(cellValue);
            int width = columnProperty.getWidth();
            if (cellValueLength > width) {
                // 设置字段单元格长度
                columnProperty.setWidth(cellValueLength);
            }
        } else {
            // 计算格式化字符串的长度
            columnProperty.setWidth(StringUtil.getValueLength(columnProperty.getFormatValue()));
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

                fillCellValue(cell, columnProperty.getCellStyle(), cellValue, columnProperty.getFormatValue());
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
                tempValue = handleResourceBundle((String) map.get(cellValue.toString()));
                break;
            }
        }
        return tempValue;
    }

    /**
     * 设置单元格值
     * 判断单元格数据类型
     * 数值型 又分为 整数 小数 （数字长度超过 15 会丢失精度）
     * 文本型 百分比数值 时间格式化 也就是字符串
     * 日期时间型 需要格式化字符串
     * @param cell
     * @param cellStyle
     * @param cellValue
     * @param formatValue 日期格式化字符串
     */
    private void fillCellValue(Cell cell, CellStyle cellStyle, Object cellValue, String formatValue) {
        // 格式化
        DataFormat dataFormat = this.workbook.createDataFormat();

        boolean isNum = DoExcelUtil.isNum(cellValue);
        boolean isInteger = DoExcelUtil.isIntger(cellValue);
        boolean isPercent = DoExcelUtil.isPercent(cellValue);
        // 整数
        if (isNum && isPercent == false) {
            // 整数不显示小数
            if (isInteger) {
                // 如果整数长度大于 15 转换成文本
                if (String.valueOf(cellValue).length() > 15) {
                    cellStyle.setDataFormat(dataFormat.getFormat("@"));
                    cell.setCellValue(cellValue.toString());
                } else {
                    cellStyle.setDataFormat(dataFormat.getFormat("0"));
                    cell.setCellValue(Double.parseDouble(cellValue.toString()));
                }
            } else {
                cellStyle.setDataFormat(dataFormat.getFormat("0.00"));
                cell.setCellValue(Double.parseDouble(cellValue.toString()));
            }
        } else {
            // 日期时间格式 JDK 1.8 之后的日期格式暂时没实现
            if (cellValue instanceof Date) {
                if (formatValue == null || formatValue.length() == 0)
                    cellStyle.setDataFormat(dataFormat.getFormat(Const.YMDHMS));
                else
                    cellStyle.setDataFormat(dataFormat.getFormat(formatValue));
                cell.setCellValue((Date) cellValue);
            } else {
                // 字符串
                cellStyle.setDataFormat(dataFormat.getFormat("@"));
                cell.setCellValue(cellValue.toString());
            }
        }
        cell.setCellStyle(cellStyle);
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
     * 主要处理表格标题和国际化资源配置
     * @param clazz
     */
    private void addSheet(Class clazz) {
        // 获取类上的注解
        DoSheet doSheet = (DoSheet) clazz.getAnnotation(DoSheet.class);

        // 设置国际化资源
        if (doSheet != null) {
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
        // 设置表标题
        String title = doSheet != null ? doSheet.title() : Const.UNTITLED;
        this.currentSheet = DoExcelUtil.createSheet(this.workbook, handleResourceBundle(title));
    }

    /**
     * 添加行数据
     * @param list 数据集合
     * @param localeFlag 是否需要国际化
     */
    public void addRow(List<Object> list, boolean localeFlag) {
        // 单元格样式
        CellStyle cellStyle = this.workbook.createCellStyle();

        Row row = getNextRow();
        // 遍历数据
        for (int i = 0; i < list.size(); i++) {
            int cellNum = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
            Cell cell = row.createCell(cellNum);

            // 非空判断
            Object cellValue = list.get(i) == null ? "" : list.get(i);

            // 是否需要国际化
            if (localeFlag && cellValue instanceof String) {
                cellValue = handleResourceBundle(cellValue.toString());
            }

            fillCellValue(cell, cellStyle, cellValue, "");
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
