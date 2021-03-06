package model;

import com.hotstrip.DoExcel.annotation.Column;
import com.hotstrip.DoExcel.annotation.ColumnEnum;
import com.hotstrip.DoExcel.annotation.DoSheet;

import java.math.BigDecimal;
import java.util.Date;

@DoSheet(title = "终端信息表", localeResource = "locales/excel_agent")
public class ExcelAgent {
    @Column(name = "agentId")
    private Long agentId;
    @Column(name = "agentName")
    private String agentName;
    @Column(name = "agentCode")
    private String agentCode;
    @Column(name = "one")
    private boolean one;
    @Column(name = "two", columnEnums = { @ColumnEnum(code = "0", value = "zero"), @ColumnEnum(code = "1", value = "value 1")})
    private Integer two;
    @Column(name = "three")
    private BigDecimal three;
    @Column(name = "createTime", format = "YYYY-MM-dd HH:mm:ss")
    private Date createTime;

    public Long getAgentId() {
        return agentId;
    }

    public void setAgentId(Long agentId) {
        this.agentId = agentId;
    }

    public String getAgentName() {
        return agentName;
    }

    public void setAgentName(String agentName) {
        this.agentName = agentName;
    }

    public String getAgentCode() {
        return agentCode;
    }

    public void setAgentCode(String agentCode) {
        this.agentCode = agentCode;
    }

    public boolean isOne() {
        return one;
    }

    public void setOne(boolean one) {
        this.one = one;
    }

    public Integer getTwo() {
        return two;
    }

    public void setTwo(Integer two) {
        this.two = two;
    }

    public BigDecimal getThree() {
        return three;
    }

    public void setThree(BigDecimal three) {
        this.three = three;
    }

    public Date getCreateTime() {
        return createTime;
    }

    public void setCreateTime(Date createTime) {
        this.createTime = createTime;
    }

    @Override
    public String toString() {
        return "ExcelAgent{" +
                "agentId=" + agentId +
                ", agentName='" + agentName + '\'' +
                ", agentCode='" + agentCode + '\'' +
                '}';
    }
}
