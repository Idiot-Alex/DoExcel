import com.hotstrip.annotation.Column;

import java.math.BigDecimal;

/**
 * Created by idiot on 2019/6/2.
 */
public class OrderSheet {
    @Column(name = "交易时间")
    private String createTime;

    @Column(name = "公众账号ID")
    private String appId;

    @Column(name = "商户号")
    private String machId;

    @Column(name = "设备号")
    private String agentCode;

    @Column(name = "微信订单号")
    private String tradeNo;

    @Column(name = "商户订单号")
    private String orderId;

    @Column(name = "用户标识")
    private String openId;

    @Column(name = "交易类型")
    private String tradeType;

    @Column(name = "交易状态")
    private String success;

    @Column(name = "应结订单金额")
    private String priceText;
    private BigDecimal price;

    @Column(name = "商品名称")
    private String goodsBody;

    public String getCreateTime() {
        return createTime;
    }

    public void setCreateTime(String createTime) {
        this.createTime = createTime;
    }

    public String getAppId() {
        return appId;
    }

    public void setAppId(String appId) {
        this.appId = appId;
    }

    public String getMachId() {
        return machId;
    }

    public void setMachId(String machId) {
        this.machId = machId;
    }

    public String getAgentCode() {
        return agentCode;
    }

    public void setAgentCode(String agentCode) {
        this.agentCode = agentCode;
    }

    public String getTradeNo() {
        return tradeNo;
    }

    public void setTradeNo(String tradeNo) {
        this.tradeNo = tradeNo;
    }

    public String getOrderId() {
        return orderId;
    }

    public void setOrderId(String orderId) {
        this.orderId = orderId;
    }

    public String getOpenId() {
        return openId;
    }

    public void setOpenId(String openId) {
        this.openId = openId;
    }

    public String getTradeType() {
        return tradeType;
    }

    public void setTradeType(String tradeType) {
        this.tradeType = tradeType;
    }

    public String getSuccess() {
        return success;
    }

    public void setSuccess(String success) {
        this.success = success;
    }

    public String getPriceText() {
        return priceText;
    }

    public void setPriceText(String priceText) {
        this.priceText = priceText;
    }

    public BigDecimal getPrice() {
        return new BigDecimal(this.priceText.substring(1));
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public String getGoodsBody() {
        return goodsBody;
    }

    public void setGoodsBody(String goodsBody) {
        this.goodsBody = goodsBody;
    }
}
