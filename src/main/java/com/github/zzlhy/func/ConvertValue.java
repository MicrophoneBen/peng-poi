package com.github.zzlhy.func;

/**
 * 字段值转换方法接口
 * Created by Administrator on 2017-11-29.
 */
public interface ConvertValue {
    /**
     * 导出时,值的转换
     * @param obj obj
     * @return 值
     */
    String convert(Object obj);

    /**
     * 导入时,值转换
     * @param obj obj
     * @return 值
     */
    String convert(String obj);
}
