package com.github.zzlhy.entity;

import org.apache.poi.hssf.util.HSSFColor;

/**
 * 对POI的颜色类的封装,简化使用
 * Created by Administrator on 2018-10-31.
 */
public enum  Color {
    BLACK                (0x08,   -1, 0x000000),
    BROWN                (0x3C,   -1, 0x993300),
    OLIVE_GREEN          (0x3B,   -1, 0x333300),
    DARK_GREEN           (0x3A,   -1, 0x003300),
    DARK_TEAL            (0x38,   -1, 0x003366),
    DARK_BLUE            (0x12, 0x20, 0x000080),
    INDIGO               (0x3E,   -1, 0x333399),
    GREY_80_PERCENT      (0x3F,   -1, 0x333333),
    ORANGE               (0x35,   -1, 0xFF6600),
    DARK_YELLOW          (0x13,   -1, 0x808000),
    GREEN                (0x11,   -1, 0x008000),
    TEAL                 (0x15, 0x26, 0x008080),
    BLUE                 (0x0C, 0x27, 0x0000FF),
    BLUE_GREY            (0x36,   -1, 0x666699),
    GREY_50_PERCENT      (0x17,   -1, 0x808080),
    RED                  (0x0A,   -1, 0xFF0000),
    LIGHT_ORANGE         (0x34,   -1, 0xFF9900),
    LIME                 (0x32,   -1, 0x99CC00),
    SEA_GREEN            (0x39,   -1, 0x339966),
    AQUA                 (0x31,   -1, 0x33CCCC),
    LIGHT_BLUE           (0x30,   -1, 0x3366FF),
    VIOLET               (0x14, 0x24, 0x800080),
    GREY_40_PERCENT      (0x37,   -1, 0x969696),
    PINK                 (0x0E, 0x21, 0xFF00FF),
    GOLD                 (0x33,   -1, 0xFFCC00),
    YELLOW               (0x0D, 0x22, 0xFFFF00),
    BRIGHT_GREEN         (0x0B,   -1, 0x00FF00),
    TURQUOISE            (0x0F, 0x23, 0x00FFFF),
    DARK_RED             (0x10, 0x25, 0x800000),
    SKY_BLUE             (0x28,   -1, 0x00CCFF),
    PLUM                 (0x3D, 0x19, 0x993366),
    GREY_25_PERCENT      (0x16,   -1, 0xC0C0C0),
    ROSE                 (0x2D,   -1, 0xFF99CC),
    LIGHT_YELLOW         (0x2B,   -1, 0xFFFF99),
    LIGHT_GREEN          (0x2A,   -1, 0xCCFFCC),
    LIGHT_TURQUOISE      (0x29, 0x1B, 0xCCFFFF),
    PALE_BLUE            (0x2C,   -1, 0x99CCFF),
    LAVENDER             (0x2E,   -1, 0xCC99FF),
    WHITE                (0x09,   -1, 0xFFFFFF),
    CORNFLOWER_BLUE      (0x18,   -1, 0x9999FF),
    LEMON_CHIFFON        (0x1A,   -1, 0xFFFFCC),
    MAROON               (0x19,   -1, 0x7F0000),
    ORCHID               (0x1C,   -1, 0x660066),
    CORAL                (0x1D,   -1, 0xFF8080),
    ROYAL_BLUE           (0x1E,   -1, 0x0066CC),
    LIGHT_CORNFLOWER_BLUE(0x1F,   -1, 0xCCCCFF),
    TAN                  (0x2F,   -1, 0xFFCC99),
    AUTOMATIC            (0x40,   -1, 0x000000);

    private HSSFColor color;

    Color(int index, int index2, int rgb) {
        this.color = new HSSFColor(index, index2, new java.awt.Color(rgb));
    }

    /**
     * @see HSSFColor#getIndex()
     */
    public short getIndex() {
        return color.getIndex();
    }

    /**
     * @see HSSFColor#getIndex2()
     */
    public short getIndex2() {
        return color.getIndex2();
    }

    /**
     * @see HSSFColor#getTriplet()
     */
    public short [] getTriplet() {
        return color.getTriplet();
    }

    /**
     * @see HSSFColor#getHexString()
     */
    public String getHexString() {
        return color.getHexString();
    }


}
