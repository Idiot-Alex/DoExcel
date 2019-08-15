import com.alibaba.fastjson.JSON;
import com.hotstrip.enums.ExcelTypeEnums;
import com.hotstrip.excel.ExcelWriter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class AppTest {

    private static Logger logger = LoggerFactory.getLogger(AppTest.class);

    public static void main(String[] args) throws FileNotFoundException {
        FileOutputStream fileOutputStream = new FileOutputStream("/Users/zhangxin/Desktop/test.xlsx");
        ExcelWriter excelWriter = new ExcelWriter(fileOutputStream, ExcelTypeEnums.XLSX);

        logger.info("excelWriter: {}", JSON.toJSONString(excelWriter));
    }
}
