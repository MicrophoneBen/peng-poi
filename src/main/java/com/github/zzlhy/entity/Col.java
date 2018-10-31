package com.github.zzlhy.entity;


import com.github.zzlhy.func.ConvertValue;

/**
 * 导出数据列设置
 * Created by Administrator on 2017-11-29.
 */
public class Col {

    //标题名称
    private String title;

    //属性名
    private String key;

    //列宽度,默认为15个字符
    private int width = 15;

    //日期格式
    private String format;

    //属性值转换其他值方法
    private ConvertValue convertValue;

    //公式或函数(如果设置了公式,则不用检测key)
    private String formula;

    //列样式
    private ColStyle colStyle=new ColStyle();

    public static Col of() {
        return new Col();
    }

    public static Col of(String key) {
        return of().setKey(key);
    }

    public static Col of(String title, String key) {
        return of(key).setTitle(title);
    }

    public static Col of(String key, ConvertValue convertValue) {
        return of(key).setConvertValue(convertValue);
    }

    public static Col of(String title, int width, String formula) {
        return of().setTitle(title).setWidth(width).setFormula(formula);
    }

    public static Col of(String title, String key, String format) {
        return of(title,key).setFormat(format);
    }

    public static Col of(String title, String key, int width) {
        return of(title,key).setWidth(width);
    }

    public static Col of(String title, String key, ConvertValue convertValue) {
        return of(title,key).setConvertValue(convertValue);
    }

    public static Col of(String title, String key, ColStyle colStyle) {
        return of(title,key).setColStyle(colStyle);
    }

    public static Col of(String title, String key, int width, String format) {
        return of(title,key).setWidth(width).setFormat(format);
    }

    public static Col of(String title, String key, int width, ColStyle colStyle) {
        return of(title,key,colStyle).setWidth(width);
    }

    public static Col of(String title, String key, int width, ConvertValue convertValue) {
        return of(title,key).setWidth(width).setConvertValue(convertValue);
    }


    public Col() {
    }

    public Col(String key) {
        this.key = key;
    }

    public Col(String title, String key) {
        this.title = title;
        this.key = key;
    }

    public Col(String key, ConvertValue convertValue) {
        this.key = key;
        this.convertValue = convertValue;
    }

    public Col(String title, int width, String formula) {
        this.title = title;
        this.width = width;
        this.formula = formula;
    }

    public Col(String title, String key, String format) {
        this.title = title;
        this.key = key;
        this.format = format;
    }

    public Col(String title, String key, int width) {
        this.title = title;
        this.key = key;
        this.width = width;
    }

    public Col(String title, String key, ConvertValue convertValue) {
        this.title = title;
        this.key = key;
        this.convertValue = convertValue;
    }

    public Col(String title, String key, ColStyle colStyle) {
        this.title = title;
        this.key = key;
        this.colStyle = colStyle;
    }

    public Col(String title, String key, int width, String format) {
        this.title = title;
        this.key = key;
        this.width = width;
        this.format = format;
    }

    public Col(String title, String key, int width, ColStyle colStyle) {
        this.title = title;
        this.key = key;
        this.width = width;
        this.colStyle = colStyle;
    }

    public Col(String title, String key, int width, ConvertValue convertValue) {
        this.title = title;
        this.key = key;
        this.width = width;
        this.convertValue = convertValue;
    }


    public String getTitle() {
        return title;
    }

    public Col setTitle(String title) {
        this.title = title;
        return this;
    }

    public String getKey() {
        return key;
    }

    public Col setKey(String key) {
        this.key = key;
        return this;
    }

    public ConvertValue getConvertValue() {
        return convertValue;
    }

    public Col setConvertValue(ConvertValue convertValue) {
        this.convertValue = convertValue;
        return this;
    }

    public int getWidth() {
        return width;
    }

    public Col setWidth(int width) {
        this.width = width;
        return this;
    }

    public String getFormat() {
        return format;
    }

    public Col setFormat(String format) {
        this.format = format;
        return this;
    }

    public String getFormula() {
        return formula;
    }

    public Col setFormula(String formula) {
        this.formula = formula;
        return this;
    }

    public ColStyle getColStyle() {
        return colStyle;
    }

    public Col setColStyle(ColStyle columnStyle) {
        this.colStyle = columnStyle;
        return this;
    }
}
