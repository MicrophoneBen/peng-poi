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

    //列样式
    private ColStyle colStyle=new ColStyle();

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

    public Col(String title, String key, int width) {
        this.title = title;
        this.key = key;
        this.width = width;
    }

    public Col(String title, String key, String format) {
        this.title = title;
        this.key = key;
        this.format = format;
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

    public Col(String title, String key, String format, ConvertValue convertValue) {
        this.title = title;
        this.key = key;
        this.format = format;
        this.convertValue = convertValue;
    }

    public Col(String title, String key, int width, ConvertValue convertValue) {
        this.title = title;
        this.key = key;
        this.width = width;
        this.convertValue = convertValue;
    }

    public Col(String title, String key, int width, String format, ConvertValue convertValue) {
        this.title = title;
        this.key = key;
        this.width = width;
        this.format = format;
        this.convertValue = convertValue;
    }

    public Col(String title, String key, int width, String format, ConvertValue convertValue, ColStyle colStyle) {
        this.title = title;
        this.key = key;
        this.width = width;
        this.format = format;
        this.convertValue = convertValue;
        this.colStyle = colStyle;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public ConvertValue getConvertValue() {
        return convertValue;
    }

    public void setConvertValue(ConvertValue convertValue) {
        this.convertValue = convertValue;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public ColStyle getColStyle() {
        return colStyle;
    }

    public void setColStyle(ColStyle columnStyle) {
        this.colStyle = columnStyle;
    }
}
