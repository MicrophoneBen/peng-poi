package com.github.zzlhy.entity;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * 标题行样式属性
 * Created by Administrator on 2017-11-29.
 */
public class HeadRowStyle {

    //标题行加粗
    private Boolean headBold=true;

    //标题行位置  默认居左
    private HorizontalAlignment horizontalAlignment= HorizontalAlignment.LEFT;

    public Boolean getHeadBold() {
        return headBold;
    }

    public void setHeadBold(Boolean headBold) {
        this.headBold = headBold;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }
}
