package model;

import com.hotstrip.annotation.DoSheet;

public class ExcelAgent {
    @DoSheet(name = "终端ID")
    private Long agentId;
    @DoSheet(name = "终端名称")
    private String agentName;
    @DoSheet(name = "终端编号")
    private String agentCode;

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
}
