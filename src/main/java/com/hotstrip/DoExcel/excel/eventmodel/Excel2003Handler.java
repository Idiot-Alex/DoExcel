package com.hotstrip.DoExcel.excel.eventmodel;

import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2019/7/29.
 * 处理 Excel 2003 版本
 * 实现 HSSFListener 监听器，事件驱动模型解析
 */
public class Excel2003Handler implements HSSFListener, ExcelHandler {
    private static Logger logger = LoggerFactory.getLogger(Excel2003Handler.class);

    private POIFSFileSystem poifsFileSystem;

    // 当前单元表下标
    private int sheetIndex = -1;

    // 当前行号
    private int currentRowNum;

    /**
     * 监听方法，处理 Record
     * 关于常用的 Record 的 sid 如下:
     * BoundSheetRecord: 记录了sheetName
     * BOFRecord: Workbook、Sheet 的开始
     * BlankRecord: 存在单元格样式的空单元格
     * BoolErrRecord: 布尔或错误单元格
     * FormulaRecord: 公式单元格
     * StringRecord: 公式的计算结果单元格
     * LabelRecord: 文本单元格
     * LabelSSTRecord: 共用的文本单元格
     * NumberRecord: 数值单元格：数字单元格和日期单元格
     * EOFRecord: Workbook、Sheet的结束
     * RowRecord: 行记录
     * @param record
     */
    public void processRecord(Record record) {
        switch (record.getSid()){
            case BoundSheetRecord.sid:
                BoundSheetRecord boundSheetRecord = (BoundSheetRecord) record;
                logger.info("current process workSheet is: [{}]", boundSheetRecord.getSheetname());
                break;
            case BOFRecord.sid:
                BOFRecord bofRecord = (BOFRecord) record;
                // 如果类型是 单元表，当前单元表下标 +1
                if (bofRecord.getType() == BOFRecord.TYPE_WORKSHEET)
                    sheetIndex ++;
                break;
            case BlankRecord.sid:
                break;
            case BoolErrRecord.sid:
                break;
            case FormulaRecord.sid:
                break;
            case StringRecord.sid:
                break;
            case LabelRecord.sid:
                break;
            case LabelSSTRecord.sid:
                break;
            case NumberRecord.sid:
                break;
            case EOFRecord.sid:
                break;
            case RowRecord.sid:
                RowRecord rowRecord = (RowRecord) record;
                logger.info("rowNumber: [{}]...firstCol: [{}]...lastCol: [{}]", rowRecord.getRowNumber(), rowRecord.getFirstCol(), rowRecord.getLastCol());
                break;
            default:
                break;
        }
    }

    /**
     * 处理单元表
     * @param fileInputStream
     */
    public void readSheet(FileInputStream fileInputStream) throws IOException {
        // 输入流
        this.poifsFileSystem = new POIFSFileSystem(fileInputStream);
        // 监听器
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        FormatTrackingHSSFListener formatListener = new FormatTrackingHSSFListener(listener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        request.addListenerForAllRecords(formatListener);
        // 处理 Excel 工作薄
        factory.processWorkbookEvents(request, this.poifsFileSystem);
    }

    public static void main(String [] args) throws IOException {
        FileInputStream fis = new FileInputStream("F:\\Excel2003.xls");
        ExcelHandler excelHandler = new Excel2003Handler();
        excelHandler.readSheet(fis);
    }
}
