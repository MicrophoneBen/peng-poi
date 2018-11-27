package com.github.zzlhy.annotation;

import com.github.zzlhy.func.ConvertValue;

import java.lang.annotation.*;

/**
 * 列属性注解
 * Created by Administrator on 2018-11-27.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Documented
public @interface Col {

    //标题名称
    String value() default "";

    //标题名称
    String title() default "";

    //属性名
    String key() default "";

    //列宽度,默认为15个字符
    int width() default 15;

    //日期格式
    String format() default "";

    //属性值转换其他值方法
    Class<? extends ConvertValue> convertValue() default ConvertValue.class;

    //公式或函数(如果设置了公式,则不用检测key)
    String formula() default "";

    //下拉列表的数组
    String[] dropdownList() default {};
}
