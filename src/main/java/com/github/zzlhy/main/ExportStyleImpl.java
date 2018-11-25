package com.github.zzlhy.main;

import com.github.zzlhy.entity.*;
import com.github.zzlhy.util.Utils;
import org.apache.poi.ss.usermodel.*;

/**
 * Created by Administrator on 2018-11-25.
 */
public class ExportStyleImpl implements ExportStyle {
    /**
     * 创建单元格样式
     *
     * @param workbook workbook
     * @param style 样式配置类
     * @return CellStyle
     */
    @Override
    public CellStyle createCellStyle(Workbook workbook, ColStyleAbstract style) {
        //单元格样式
        CellStyle cellStyle = null;

        if(workbook == null){
            return null;
        }

        if(style == null){
            return null;
        }

        FontStyle fontStyle = style.getFontStyle();
        HorizontalAlignment horizontalAlignment = style.getHorizontalAlignment();
        VerticalAlignment verticalAlignment = style.getVerticalAlignment();
        if(fontStyle != null || horizontalAlignment != null || verticalAlignment != null ||
                style.getBackgroundColor() != null || style.getTopBorder() != null|| style.getBottomBorder() != null||
                style.getLeftBorder() != null|| style.getRightBorder() != null){
            cellStyle = workbook.createCellStyle();
        }

        if(fontStyle != null){
            //字体样式
            Font font = workbook.createFont();
            if(fontStyle.getBold() != null){
                font.setBold(fontStyle.getBold());
            }
            if(fontStyle.getItalic() != null){
                font.setItalic(fontStyle.getItalic());
            }
            if(fontStyle.getStrikeout() != null){
                font.setStrikeout(fontStyle.getStrikeout());
            }
            if(fontStyle.getColor() != null){
                font.setColor(fontStyle.getColor());
            }
            if(fontStyle.getHeightInPoints() != null){
                font.setFontHeightInPoints(fontStyle.getHeightInPoints());
            }
            if(fontStyle.getUnderline() != null){
                font.setUnderline(fontStyle.getUnderline());
            }
            if(fontStyle.getTypeOffset() != null){
                font.setTypeOffset(fontStyle.getTypeOffset());
            }
            if(Utils.notEmpty(fontStyle.getFontName())){
                font.setFontName(fontStyle.getFontName());
            }
            cellStyle.setFont(font);
        }
        if(horizontalAlignment != null){
            cellStyle.setAlignment(horizontalAlignment);
        }
        if(verticalAlignment != null){
            cellStyle.setVerticalAlignment(verticalAlignment);
        }
        //背景色
        if(style.getBackgroundColor() != null){
            cellStyle.setFillForegroundColor(style.getBackgroundColor());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        //边框
        if(style.getTopBorder() != null){
            cellStyle.setBorderTop(style.getTopBorder());
            if(style.getTopBorderColor() != null){
                cellStyle.setTopBorderColor(style.getTopBorderColor());
            }
        }
        if(style.getBottomBorder() != null){
            cellStyle.setBorderBottom(style.getBottomBorder());
            if(style.getBottomBorderColor() != null){
                cellStyle.setBottomBorderColor(style.getBottomBorderColor());
            }
        }
        if(style.getLeftBorder() != null){
            cellStyle.setBorderLeft(style.getLeftBorder());
            if(style.getLeftBorderColor() != null){
                cellStyle.setLeftBorderColor(style.getLeftBorderColor());
            }
        }
        if(style.getRightBorder() != null){
            cellStyle.setBorderRight(style.getRightBorder());
            if(style.getRightBorderColor() != null){
                cellStyle.setRightBorderColor(style.getRightBorderColor());
            }
        }

        return cellStyle;
    }
}
