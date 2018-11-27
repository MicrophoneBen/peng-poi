package com.github.zzlhy.annotation;

import com.github.zzlhy.entity.ExcelType;
import java.lang.annotation.*;

/**
 * Created by Administrator on 2018-11-27.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {

    //Sheet名称
    String sheetName() default "Sheet";

    //每个sheet允许的数据行数,超过此行数会创建新的sheet (xlsx最大值为1048576)
    int sheetDataTotal() default 500000;

    //导出时开始写入的起始行,默认从0开始
    int writeRow() default 0;

    //导入时,起始读取行
    int readRow() default 1;

    //行高度
    float height() default 15;

    //是否创建标题行
    boolean createHeadRow() default true;

    //导出Excel类型xls/xlsx   默认为xlsx
    ExcelType excelType() default ExcelType.XLSX;

    //冻结列设置
    //固定行的列号 是从1开始
    int freezeColSplit() default 0;
    //固定行的行号
    int freezeRowSplit() default 0;


}
