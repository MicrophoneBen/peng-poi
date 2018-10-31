package com.github.zzlhy.entity;

/**
 * 字体设置
 * Created by Administrator on 2018-10-31.
 */
public class FontStyle {

    //字体名称
    private String fontName;

    //字体颜色
    private Short color;

    //字体字号
    private Short heightInPoints;

    //设置斜体
    private Boolean italic;

    //设置加粗
    private Boolean bold;

    //设置下划线
    //单下划线 FontFormatting.U_SINGLE      1
    //双下划线 FontFormatting.U_DOUBLE      2
    //会计用单下划线 FontFormatting.U_SINGLE_ACCOUNTING    3
    //会计用双下划线 FontFormatting.U_DOUBLE_ACCOUNTING    4
    //无下划线 FontFormatting.U_NONE        0
    private Byte underline;

    //设置删除线
    private Boolean strikeout;

    //设置上标、下标
    //上标 FontFormatting.SS_SUPER
    //下标 FontFormatting.SS_SUB
    //普通，默认值 FontFormatting.SS_NONE
    private Short typeOffset;

    public static FontStyle of(){
        return new FontStyle();
    }

    public static FontStyle of(Short color){
        return of().setColor(color);
    }

    public String getFontName() {
        return fontName;
    }

    public FontStyle setFontName(String fontName) {
        this.fontName = fontName;
        return this;
    }

    public Short getColor() {
        return color;
    }

    public FontStyle setColor(Short color) {
        this.color = color;
        return this;
    }

    public Short getHeightInPoints() {
        return heightInPoints;
    }

    public FontStyle setHeightInPoints(Short heightInPoints) {
        this.heightInPoints = heightInPoints;
        return this;
    }

    public Boolean getItalic() {
        return italic;
    }

    public FontStyle setItalic(Boolean italic) {
        this.italic = italic;
        return this;
    }

    public Boolean getBold() {
        return bold;
    }

    public FontStyle setBold(Boolean bold) {
        this.bold = bold;
        return this;
    }

    public Byte getUnderline() {
        return underline;
    }

    public FontStyle setUnderline(Byte underline) {
        this.underline = underline;
        return this;
    }

    public Boolean getStrikeout() {
        return strikeout;
    }

    public FontStyle setStrikeout(Boolean strikeout) {
        this.strikeout = strikeout;
        return this;
    }

    public Short getTypeOffset() {
        return typeOffset;
    }

    public FontStyle setTypeOffset(Short typeOffset) {
        this.typeOffset = typeOffset;
        return this;
    }
}
