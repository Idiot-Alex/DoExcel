import com.hotstrip.DoExcel.enums.ExcelTypeEnums;
import com.hotstrip.DoExcel.excel.ExcelWriter;
import model.ExcelAgent;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Export Excel Tests
 */
public class ExportTest {
    private static Logger logger = LoggerFactory.getLogger(ExportTest.class);

    private List<ExcelAgent> list = new ArrayList<ExcelAgent>();

    // 构建测试数据
    @Before
    public void initData(){
        for (int i=0; i<10; i++){
            ExcelAgent excelAgent = new ExcelAgent();
            excelAgent.setAgentCode("code" + i);
            excelAgent.setAgentName("name" + i);
            excelAgent.setAgentId(i * 10L);
            excelAgent.setTwo(i * 10);
            excelAgent.setThree(BigDecimal.valueOf(i).multiply(BigDecimal.valueOf(10.1)));
            excelAgent.setCreateTime(new Date());
            list.add(excelAgent);
        }
    }

    // 简单使用
    @Test
    public void quickStart() throws FileNotFoundException {
        // FileOutputStream fileOutputStream = new FileOutputStream("F:\\test.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/zhangxin/Desktop/test.xlsx");
        ExcelWriter excelWriter = new ExcelWriter(fileOutputStream, ExcelTypeEnums.XLSX);

        excelWriter.write(list, ExcelAgent.class)
                .close();
    }
}
