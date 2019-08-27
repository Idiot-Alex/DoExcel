package com.hotstrip;

import com.hotstrip.annotation.Column;

import java.math.BigDecimal;

/**
 * Created by idiot on 2019/6/2.
 */
public class SeestechOrderSheet {
    @Column(name = "订单id")
    private String orderId;

    @Column(name = "所属企业")
    private String companyName;

    @Column(name = "终端名称")
    private String agentName;

    @Column(name = "终端编号")
    private String agentCode;

    @Column(name = "商品名称")
    private String goodsBody;

    @Column(name = "商品code")
    private String goodsCode;

    @Column(name = "交易金额")
    private String priceText;
    private BigDecimal price;

    @Column(name = "支付状态")
    private String successText;

    @Column(name = "支付渠道")
    private String channelText;

    @Column(name = "出货状态")
    private String saleStatusText;

    @Column(name = "创建时间")
    private String createTime;

    public String getOrderId() {
        return orderId;
    }

    public void setOrderId(String orderId) {
        this.orderId = orderId;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
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

    public String getGoodsBody() {
        return goodsBody;
    }

    public void setGoodsBody(String goodsBody) {
        this.goodsBody = goodsBody;
    }

    public String getGoodsCode() {
        return goodsCode;
    }

    public void setGoodsCode(String goodsCode) {
        this.goodsCode = goodsCode;
    }

    public String getPriceText() {
        return priceText;
    }

    public void setPriceText(String priceText) {
        this.priceText = priceText;
    }

    public BigDecimal getPrice() {
        return new BigDecimal(this.priceText);
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public String getSuccessText() {
        return successText;
    }

    public void setSuccessText(String successText) {
        this.successText = successText;
    }

    public String getChannelText() {
        return channelText;
    }

    public void setChannelText(String channelText) {
        this.channelText = channelText;
    }

    public String getSaleStatusText() {
        return saleStatusText;
    }

    public void setSaleStatusText(String saleStatusText) {
        this.saleStatusText = saleStatusText;
    }

    public String getCreateTime() {
        return createTime;
    }

    public void setCreateTime(String createTime) {
        this.createTime = createTime;
    }
}
