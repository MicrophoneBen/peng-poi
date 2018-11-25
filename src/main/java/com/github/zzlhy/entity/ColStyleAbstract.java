package com.github.zzlhy.entity;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 样式配置抽象类
 * Created by Administrator on 2018-11-25.
 */
public abstract class ColStyleAbstract {

    //单元格水平对齐方式
    private HorizontalAlignment horizontalAlignment;

    //单元格垂直对齐方式
    private VerticalAlignment verticalAlignment;

    //单元格背景颜色
    private Short backgroundColor;

    //字体设置
    private FontStyle fontStyle;

    //边框样式
    private BorderStyle topBorder;

    private BorderStyle bottomBorder;

    private BorderStyle leftBorder;

    private BorderStyle rightBorder;

    //边框颜色
    private Short topBorderColor;

    private Short bottomBorderColor;

    private Short leftBorderColor;

    private Short rightBorderColor;

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public ColStyleAbstract setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;

    }

    public ColStyleAbstract setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    public Short getBackgroundColor() {
        return backgroundColor;
    }

    public ColStyleAbstract setBackgroundColor(Short backgroundColor) {
        this.backgroundColor = backgroundColor;
        return this;
    }

    public FontStyle getFontStyle() {
        return fontStyle;
    }

    public ColStyleAbstract setFontStyle(FontStyle fontStyle) {
        this.fontStyle = fontStyle;
        return this;
    }

    public BorderStyle getTopBorder() {
        return topBorder;
    }

    public ColStyleAbstract setTopBorder(BorderStyle topBorder) {
        this.topBorder = topBorder;
        return this;
    }

    public BorderStyle getBottomBorder() {
        return bottomBorder;
    }

    public ColStyleAbstract setBottomBorder(BorderStyle bottomBorder) {
        this.bottomBorder = bottomBorder;
        return this;
    }

    public BorderStyle getLeftBorder() {
        return leftBorder;
    }

    public ColStyleAbstract setLeftBorder(BorderStyle leftBorder) {
        this.leftBorder = leftBorder;
        return this;
    }

    public BorderStyle getRightBorder() {
        return rightBorder;
    }

    public ColStyleAbstract setRightBorder(BorderStyle rightBorder) {
        this.rightBorder = rightBorder;
        return this;
    }

    public Short getTopBorderColor() {
        return topBorderColor;
    }

    public ColStyleAbstract setTopBorderColor(Short topBorderColor) {
        this.topBorderColor = topBorderColor;
        return this;
    }

    public Short getBottomBorderColor() {
        return bottomBorderColor;
    }

    public ColStyleAbstract setBottomBorderColor(Short bottomBorderColor) {
        this.bottomBorderColor = bottomBorderColor;
        return this;
    }

    public Short getLeftBorderColor() {
        return leftBorderColor;
    }

    public ColStyleAbstract setLeftBorderColor(Short leftBorderColor) {
        this.leftBorderColor = leftBorderColor;
        return this;
    }

    public Short getRightBorderColor() {
        return rightBorderColor;
    }

    public ColStyleAbstract setRightBorderColor(Short rightBorderColor) {
        this.rightBorderColor = rightBorderColor;
        return this;
    }
}
