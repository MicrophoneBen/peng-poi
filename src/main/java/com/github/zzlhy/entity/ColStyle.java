package com.github.zzlhy.entity;

/**
 * 列样式属性
 * Created by Administrator on 2017-11-29.
 */
public class ColStyle {

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
}
