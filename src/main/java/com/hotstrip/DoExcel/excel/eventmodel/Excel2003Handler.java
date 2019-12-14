package com.hotstrip.DoExcel.excel.eventmodel;

import com.alibaba.fastjson.JSON;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
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

    // 当前行号 0 表示第一行
    private int currentRowNum;
    // 当前列号 0 表示第一列
    private int currentColNum;

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
     * FormulaRecord: 公式单元格 => 比如 =SUM(B1:B2)
     * StringRecord: 公式的计算结果单元格 => 比如 =(B2)
     * LabelSSTRecord: 共用的文本单元格
     * NumberRecord: 数值单元格：数字单元格和日期单元格
     * EOFRecord: Workbook、Sheet的结束
     * @param record
     */
    public void processRecord(Record record) {
        switch (record.getSid()){
            case BoundSheetRecord.sid:
                BoundSheetRecord boundSheetRecord = (BoundSheetRecord) record;
                logger.debug("current process workSheet is: [{}]", boundSheetRecord.getSheetname());
                break;
            case BOFRecord.sid:
                BOFRecord bofRecord = (BOFRecord) record;
                // 如果类型是 单元表，当前单元表下标 +1
                if (bofRecord.getType() == BOFRecord.TYPE_WORKSHEET){
                    this.sheetIndex ++;
                    // 当前行号重置为 0
                    this.currentRowNum = 0;
                }
                break;
            case BlankRecord.sid:
                // BlankRecord blankRecord = (BlankRecord) record;
                addFieldValue("");
                break;
            case BoolErrRecord.sid:
                BoolErrRecord boolErrRecord = (BoolErrRecord) record;
                String value = boolErrRecord.isBoolean() ? String.valueOf(boolErrRecord.getBooleanValue()) : String.valueOf(boolErrRecord.getErrorValue());
                addFieldValue(value);
                break;
            case FormulaRecord.sid:
                FormulaRecord formulaRecord = (FormulaRecord) record;
                /**
                 * 该判断用于区分 =(B3) 和 =SUM(B1:B4)
                 * 换句话说，区分赋值类公式和函数类公式（并未做完整测试，可能不准确）
                 * ----------------------------------------------------------------------
                 * 该判断结果若为 true 就会触发 FormulaRecord.sid 和 StringRecord.sid 两种 sid
                 * 但是实际上只有在 StringRecord.sid 里面才能取到真正的值
                 */
                if (!formulaRecord.hasCachedResultString()) {
                    addFieldValue(String.valueOf(formulaRecord.getValue()));
                }
                break;
            case StringRecord.sid:
                StringRecord stringRecord = (StringRecord) record;
                addFieldValue(stringRecord.getString());
                break;
            // SSTRecord 保存 xls 文件中唯一的 String
            case SSTRecord.sid:
                this.sstRecord = (SSTRecord) record;
                break;
            // LabelSSTRecord 取出索引值，集合上面的 SSTRecord 去取值
            case LabelSSTRecord.sid:
                LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                addFieldValue(sstRecord.getString(labelSSTRecord.getSSTIndex()).toString().trim());
                break;
            case NumberRecord.sid:
                NumberRecord numberRecord = (NumberRecord) record;
                addFieldValue(String.valueOf(numberRecord.getValue()));
                break;
            default:
                break;
        }
        // 空值的操作
        if (record instanceof MissingCellDummyRecord) {
            // MissingCellDummyRecord missingCellDummyRecord = (MissingCellDummyRecord) record;
            addFieldValue("");
        }
        // 如果是一行的最后一列
        if (record instanceof LastCellOfRowDummyRecord) {
            // LastCellOfRowDummyRecord lastCellOfRowDummyRecord = (LastCellOfRowDummyRecord) record;
            logger.info(JSON.toJSONString(this.rowList));
            this.rowList.clear();
            // 行号自增
            this.currentRowNum++;
            // 列号设置为 0
            this.currentColNum = 0;
        }
    }

    // 添加字段值 修改当前列号
    private void addFieldValue(String value) {
        this.rowList.add(this.currentColNum, value);
        // 列号自增
        this.currentColNum++;
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
