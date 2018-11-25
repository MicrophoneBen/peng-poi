package com.github.zzlhy.main;

import com.github.zzlhy.entity.ColStyleAbstract;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 导出样式
 * Created by Administrator on 2018-11-25.
 */
@FunctionalInterface
public interface ExportStyle {

    /**
     * 创建单元格样式
     * @param workbook workbook
     * @param styleAbstract 样式配置类
     * @return CellStyle
     */
    CellStyle createCellStyle(Workbook workbook, ColStyleAbstract styleAbstract);
}
