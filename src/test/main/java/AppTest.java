import com.alibaba.fastjson.JSON;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.ExcelWriter;
import model.ExcelAgent;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

public class AppTest {

    private static Logger logger = LoggerFactory.getLogger(AppTest.class);

    public static void main(String[] args) throws FileNotFoundException {

        List<ExcelAgent> list = new ArrayList<ExcelAgent>();
        for (int i = 0; i< 100; i++) {
            ExcelAgent excelAgent = new ExcelAgent();
            excelAgent.setAgentCode("code" + i * 9999);
            excelAgent.setAgentName("name" + i * 10000);
            excelAgent.setAgentId(i * 1000L);
            excelAgent.setTwo(i * 1.1);
            excelAgent.setThree(BigDecimal.valueOf(i).multiply(BigDecimal.valueOf(10.1)));
            list.add(excelAgent);
        }

        long startTime = System.currentTimeMillis();

        FileOutputStream fileOutputStream = new FileOutputStream("/Users/zhangxin/Desktop/test.xlsx");
        ExcelWriter excelWriter = new ExcelWriter(fileOutputStream, ExcelTypeEnums.XLSX);

        excelWriter.write(list, ExcelAgent.class)
                .close();

        long endTime = System.currentTimeMillis();

        logger.info("花费时间: {} ms", endTime - startTime);
    }
}
