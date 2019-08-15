import com.alibaba.fastjson.JSON;
import com.hotstrip.excel.ExcelWriter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class AppTest {

    private static Logger logger = LoggerFactory.getLogger(AppTest.class);

    public static void main(String[] args) {
        ExcelWriter excelWriter = ExcelWriter.builder()
                .title("test.xlsx")
                .build();
        logger.info("excelWriter: {}", JSON.toJSONString(excelWriter));
    }
}
