package com.github.zzlhy.func.impl;


import com.github.zzlhy.func.ConvertValue;

/**
 * 布尔类型的是否转换
 * Created by Administrator on 2017-11-29.
 */
public class ConvertValueBoolean implements ConvertValue {


    @Override
    public String convert(Object obj) {
        if(obj.equals(true)){
            return "是";
        }
        return "否";
    }
    public String convert(String obj) {
        if("是".equals(obj)){
            return "true";
        }
        return "false";
    }
}
