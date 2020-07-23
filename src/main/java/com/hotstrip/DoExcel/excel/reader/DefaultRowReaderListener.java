package com.hotstrip.DoExcel.excel.reader;

import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * 对 Excel 行数据读取监听的默认实现
 */
public class DefaultRowReaderListener implements IRowReaderListener {
    private static Logger logger = LoggerFactory.getLogger(DefaultRowReaderListener.class);

    /**
     * 处理 Excel 行数据
     * @param sheetIndex sheet 索引
     * @param rowIndex 行索引
     * @param rowData 数据行
     */
    public void getRow(int sheetIndex, int rowIndex, List<String> rowData) {
        logger.info("sheetIndex: {}...rowIndex: {}", sheetIndex, rowIndex);
        logger.info("rowData: {}", JSON.toJSONString(rowData));
    }
}
