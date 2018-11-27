package com.github.zzlhy.entity;

/**
 * 列样式属性
 * Created by Administrator on 2017-11-29.
 */
public class ColStyle extends ColStyleAbstract{

    public static ColStyle of() {
        return new ColStyle();
    }

    public static ColStyleAbstract of(FontStyle fontStyle) {
        return of().setFontStyle(fontStyle);
    }

    public static ColStyleAbstract of(short index){
        return of(FontStyle.of(index));
    }

}
