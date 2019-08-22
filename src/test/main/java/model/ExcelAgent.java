package model;

import com.hotstrip.annotation.Column;
import com.hotstrip.annotation.DoSheet;

import java.math.BigDecimal;
import java.util.Date;

@DoSheet(title = "终端信息表")
public class ExcelAgent {
    @Column(name = "终端ID")
    private Long agentId;
    @Column(name = "名称")
    private String agentName;
    @Column(name = "终端编号")
    private String agentCode;
    @Column(name = "one")
    private boolean one;
    @Column(name = "two")
    private Integer two;
    @Column(name = "three")
    private BigDecimal three;
    @Column(name = "创建时间")
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
