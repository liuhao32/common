package com.example.demo.annotation;

import lombok.Getter;

/**
 * Copyright (C), 2020
 *
 * @author: liuhao
 * @date: 2020/11/26 22:19
 * @description:
 */
@Getter
public class RegExpStyle {


    /**
     * 验证邮箱
     */
    public static final String EMAIL = "^([\\w-\\.]+)@((\\[[0-9]{1,3}\\.[0-9]{1,3}\\.[0-9]{1,3}\\.)|(([\\w-]+\\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$";

    /**
     * 验证网址Url
     */
    public static final String URL = "http(s)?://([\\w-]+\\.)+[\\w-]+(/[\\w- ./?%&=]*)?";

    /**
     * 验证电话号码
     */
    public static final String TELEPHONE = "^(\\d{3,4}-)?\\d{6,8}$";

    /**
     * 验证输入密码条件(字符与数据同时出现)
     */
    public static final String PASSWORD = "[A-Za-z]+[0-9]";

    /**
     * 验证输入密码长度 (6-18位)
     */
    public static final String PWDLENGTH = "^\\d{6,18}$";

    /**
     * 验证输入邮政编号
     */
    public static final String POSTALCODE = "^\\d{6}$";

    /**
     * 验证输入手机号码
     */
    public static final String HANDSET = "^^[1]+[3,5]+\\d{9}$";

    /**
     * 验证输入身份证号
     */
    public static final String IDCARD = "(^\\d{18}$)|(^\\d{15}$)";

    /**
     * 验证输入两位小数
     */
    public static final String DECIMAL = "^[0-9]+(.[0-9]{2})?$";

    /**
     * 验证输入一年的12个月
     */
    public static final String MONTH = "^(0?[[1-9]|1[0-2])$";

    /**
     * 验证输入一个月的31天
     */
    public static final String DAY = "^((0?[1-9])|((1|2)[0-9])|30|31)$";

    /**
     * 验证日期时间
     *
     * // 严格验证时间格式的(匹配[2002-01-31], [1997-04-30], // [2004-01-01])不匹配([2002-01-32], [2003-02-29], [04-01-01]) // String
     * regex = // "^((((19|20)(([02468][048])|([13579][26]))-02-29))|((20[0-9][0-9])|(19[0-9][0-9]))-((((0[1-9])|(1[0-2]))-((0[1-9])|(1\\d)|(2[0-8])))|((((0[13578])|(1[02]))-31)|(((01,3-9])|(1[0-2]))-(29|30)))))$";
     * // 没加时间验证的YYYY-MM-DD // String regex = // "^((((1[6-9]|[2-9]\\d)\\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\\d|3[01]))|(((1[6-9]|[2-9]\\d)\\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\\d|30))|(((1[6-9]|[2-9]\\d)\\d{2})-0?2-(0?[1-9]|1\\d|2[0-8]))|(((1[6-9]|[2-9]\\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-))$";
     * // 加了时间验证的YYYY-MM-DD 00:00:00
     */
    public static final String DATE = "^((((1[6-9]|[2-9]\\d)\\d{2})-(0?[13578]|1[02])-(0?[1-9]|[12]\\d|3[01]))|(((1[6-9]|[2-9]\\d)\\d{2})-(0?[13456789]|1[012])-(0?[1-9]|[12]\\d|30))|(((1[6-9]|[2-9]\\d)\\d{2})-0?2-(0?[1-9]|1\\d|2[0-8]))|(((1[6-9]|[2-9]\\d)(0[48]|[2468][048]|[13579][26])|((16|[2468][048]|[3579][26])00))-0?2-29-)) (20|21|22|23|[0-1]?\\d):[0-5]?\\d:[0-5]?\\d$";

    /**
     * 验证数字输入
     */
    public static final String NUMBER = "^[0-9]*$";

    /**
     * 验证非零的正整数
     */
    public static final String INTNUMBER = "^\\+?[1-9][0-9]*$";

    /**
     * 验证大写字母
     */
    public static final String UPCHAR = "^[A-Z]+$";

    /**
     * 验证小写字母
     */
    public static final String LOWCHAR = "^[a-z]+$";

    /**
     * 验证验证输入字母
     */
    public static final String LETTER = "^[A-Za-z]+$";

    /**
     * 验证验证输入汉字
     */
    public static final String CHINESE = "^[\u4e00-\u9fa5],{0,}$";

    /**
     * 验证输入字符串
     */
    public static final String LENGTH = "^.{8,}$";

    /**
     * 无需校验
     */
    public static final String NONE = "";

}