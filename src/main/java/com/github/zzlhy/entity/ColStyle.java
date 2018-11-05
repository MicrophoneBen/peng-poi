package com.github.zzlhy.entity;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 列样式属性
 * Created by Administrator on 2017-11-29.
 */
public class ColStyle {

    //单元格水平对齐方式
    private HorizontalAlignment horizontalAlignment;

    //单元格垂直对齐方式
    private VerticalAlignment verticalAlignment;

    //字体设置
    private FontStyle fontStyle;


    public static ColStyle of() {
        return new ColStyle();
    }

    public static ColStyle of(FontStyle fontStyle) {
        return of().setFontStyle(fontStyle);
    }

    public static ColStyle of(short index){
        return of(FontStyle.of(index));
    }

    public FontStyle getFontStyle() {
        return fontStyle;
    }

    public ColStyle setFontStyle(FontStyle fontStyle) {
        this.fontStyle = fontStyle;
        return this;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public ColStyle setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public ColStyle setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }
}
