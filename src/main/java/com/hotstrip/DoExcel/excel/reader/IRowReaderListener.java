package com.hotstrip.DoExcel.excel.reader;

import java.util.List;

/**
 * Created by Administrator on 2019/12/16.
 * 对读取 Excel 的抽象
 */
public interface IRowReaderListener {

    /**
     * 读取 Excel 行数据
     * @param sheetIndex
     * @param rowIndex
     * @param rowData
     */
    void getRow(int sheetIndex, int rowIndex, List<String> rowData);
}
