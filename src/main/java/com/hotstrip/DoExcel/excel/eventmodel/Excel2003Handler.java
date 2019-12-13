package com.hotstrip.DoExcel.excel.eventmodel;

import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingRowDummyRecord;
import org.apache.poi.hssf.record.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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

    // SSTRecord
    private SSTRecord sstRecord;

    // 当前行数据集
    private List<String> rowList = new ArrayList<String>();

    /**
     * 监听方法，处理 Record
     * 关于常用的 Record 的 sid 如下:
     * BoundSheetRecord: 记录了sheetName
     * BOFRecord: Workbook、Sheet 的开始
     * BlankRecord: 存在单元格样式的空单元格
     * BoolErrRecord: 布尔或错误单元格
     * StringRecord: 公式的计算结果单元格
     * LabelSSTRecord: 共用的文本单元格
     * NumberRecord: 数值单元格：数字单元格和日期单元格
     * EOFRecord: Workbook、Sheet的结束
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
                BlankRecord blankRecord = (BlankRecord) record;
                logger.info("blankRecord: {}", "有样式，值为空");
                break;
            case BoolErrRecord.sid:
                BoolErrRecord boolErrRecord = (BoolErrRecord) record;
                logger.info("boolErrRecord: errorValue: {} boolValue: {}", boolErrRecord.getErrorValue(), boolErrRecord.getBooleanValue());
                logger.info("boolErrRecord: isError:{} isBool: {}", boolErrRecord.isError(), boolErrRecord.isBoolean());
                break;
            case StringRecord.sid:
                StringRecord stringRecord = (StringRecord) record;
                logger.info("stringRecord: {}", stringRecord.getString());
                break;
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                logger.info("labelSSTRecord: {}", sstRecord.getString(labelSSTRecord.getSSTIndex()));
                break;
            case NumberRecord.sid:
                NumberRecord numberRecord = (NumberRecord) record;
                logger.info("numberRecord: {}", numberRecord.getValue());
                break;
            default:
                break;
        }
        // 空值的操作
        if (record instanceof MissingCellDummyRecord) {
            logger.info("该列为空值");
        }
        // 如果是一行的最后一列
        if (record instanceof LastCellOfRowDummyRecord) {
            LastCellOfRowDummyRecord lastCellOfRowDummyRecord = (LastCellOfRowDummyRecord) record;
            logger.info("lastCellOfRowDummyRecord: lastCol: {} 新行", lastCellOfRowDummyRecord.getLastColumnNumber());
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
