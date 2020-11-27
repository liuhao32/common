package com.example.demo.annotation;

import lombok.Getter;
import lombok.Setter;

/**
 * @author YCKJ3275
 */
@Getter
@Setter
public class Device {

    private String id;
    /**
     * 设备编号
     */
    @Excel(name = "deviceCode", column = 0, isRequire = true, cause = "设备编号必填")
    private String deviceCode;
    /**
     * 产品型号
     */
    @Excel(name = "productModel", column = 1, isRequire = true, cause = "产品型号必填")
    private String productModel;
    /**
     * 产品货号
     */
    @Excel(name = "productCode", column = 2, isRequire = true, cause = "产品货号必填")
    private String productCode;
    /**
     * 产品名称
     */
    @Excel(name = "productName", column = 3, isRequire = true, cause = "产品名称必填")
    private String productName;
    /**
     * 版本信息
     */
    @Excel(name = "productVersion", column = 4)
    private String productVersion;
    /**
     * 订单号
     */
    @Excel(name = "orderNo", column = 5)
    private String orderNo;
    /**
     * 客户信息
     */
    @Excel(name = "customerInfo", column = 6)
    private String customerInfo;
    /**
     * 发货地址
     */
    @Excel(name = "shipAddress", column = 7)
    private String shipAddress;
    /**
     * 发货时间
     */
    @Excel(name = "shipTime", column = 8, pattern = RegExpStyle.DATE, cause = "时间格式错误")
    private String shipTime;

}
