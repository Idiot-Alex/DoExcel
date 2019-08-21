import com.alibaba.fastjson.JSON;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.ExcelWriter;
import model.ExcelAgent;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class AppTest {

    private static Logger logger = LoggerFactory.getLogger(AppTest.class);

    public static void main(String[] args) throws FileNotFoundException {

        List<ExcelAgent> list = new ArrayList<ExcelAgent>();
        for (int i = 0; i< 100; i++) {
            ExcelAgent excelAgent = new ExcelAgent();
            excelAgent.setAgentCode("code" + i);
            excelAgent.setAgentName("name" + i);
            excelAgent.setAgentId(i * 1L);
            list.add(excelAgent);
        }

        FileOutputStream fileOutputStream = new FileOutputStream("/Users/zhangxin/Desktop/test.xlsx");
        ExcelWriter excelWriter = new ExcelWriter(fileOutputStream, ExcelTypeEnums.XLSX);

        excelWriter.write(list, ExcelAgent.class)
                .close();

        logger.info("excelWriter: {}", JSON.toJSONString(excelWriter));
    }
}
