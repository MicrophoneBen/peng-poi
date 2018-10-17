package com.github.zzlhy.func.impl;


import com.github.zzlhy.func.ConvertValue;

/**
 * Short类型的是否转换
 * Created by Administrator on 2017-11-29.
 */
public class ConvertValueShort implements ConvertValue {
    @Override
    public String convert(Object obj) {
        String s = obj.toString();
        if(s.equals("1")){
            return "是";
        }
        return "否";
    }

    @Override
    public String convert(String obj) {
        if("是".equals(obj)){
            return "1";
        }
        return "0";
    }
}
