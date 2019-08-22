package model;

import com.hotstrip.annotation.Column;
import com.hotstrip.annotation.DoSheet;

import java.math.BigDecimal;

@DoSheet(title = "终端信息表")
public class ExcelAgent {
    @Column(name = "终端ID", width = 8)
    private Long agentId;
    @Column(name = "终端名称")
    private String agentName;
    @Column(name = "终端编号")
    private String agentCode;
    @Column(name = "one")
    private boolean one;
    @Column(name = "two")
    private Double two;
    @Column(name = "three")
    private BigDecimal three;

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

    public Double getTwo() {
        return two;
    }

    public void setTwo(Double two) {
        this.two = two;
    }

    public BigDecimal getThree() {
        return three;
    }

    public void setThree(BigDecimal three) {
        this.three = three;
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
