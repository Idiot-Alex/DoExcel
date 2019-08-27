package example;

import com.hotstrip.ExcelExport;
import com.hotstrip.enums.ExcelTypeEnums;
import model.ExcelAgent;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class TestExcelAgent {

    public static void main(String[] args) throws FileNotFoundException {
        List<ExcelAgent> list = new ArrayList<ExcelAgent>();
        for (int i = 0; i < 50; i++) {
            ExcelAgent excelAgent = new ExcelAgent();
            excelAgent.setAgentId(i * 1L);
            excelAgent.setAgentName("测试终端" + i);
            excelAgent.setAgentCode("code" + i);
            list.add(excelAgent);
        }
        FileOutputStream fos = new FileOutputStream("/Users/zhangxin/Documents/test.xlsx");
        ExcelExport.exportExcel("测试", ExcelAgent.class, list, fos, ExcelTypeEnums.XLSX);
    }
}
