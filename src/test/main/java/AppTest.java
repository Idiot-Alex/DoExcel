import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.enums.LocaleEnums;
import com.hotstrip.excel.ExcelWriter;
import model.ExcelAgent;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.*;

public class AppTest {

    private static Logger logger = LoggerFactory.getLogger(AppTest.class);

    public static void main(String[] args) throws FileNotFoundException {

        List<ExcelAgent> list = new ArrayList<ExcelAgent>();
        for (int i = 0; i< 1000; i++) {
            ExcelAgent excelAgent = new ExcelAgent();
            excelAgent.setAgentCode("code" + i);
            excelAgent.setAgentName("name" + i);
            excelAgent.setAgentId(i * 10L);
            excelAgent.setTwo(i * 10);
            excelAgent.setThree(BigDecimal.valueOf(i).multiply(BigDecimal.valueOf(10.1)));
            excelAgent.setCreateTime(new Date());
            list.add(excelAgent);
        }

        long startTime = System.currentTimeMillis();

        List<Object> row1 = new ArrayList<Object>();
        List<Object> row2 = new ArrayList<Object>();
        int total = 0;
        BigDecimal total1 = BigDecimal.valueOf(0);
        for (ExcelAgent item : list) {
            total += item.getTwo();
            total1 = total1.add(item.getThree());
        }

        row1.add("整数总计");
        row1.add(total);

        row2.add("小数总计");
        row2.add(total1);

        Locale locale = LocaleEnums.getLocaleByValue("zh_cn");

        FileOutputStream fileOutputStream = new FileOutputStream("/Users/zhangxin/Desktop/test.xlsx");
        ExcelWriter excelWriter = new ExcelWriter(fileOutputStream, ExcelTypeEnums.XLSX);

        excelWriter.locale(locale)
                .write(list, ExcelAgent.class)
                .writeRow(new ArrayList())
                .writeRow(row1)
                .writeRow(row2)
                .close();

        long endTime = System.currentTimeMillis();

        logger.info("花费时间: {} ms", endTime - startTime);
    }
}
