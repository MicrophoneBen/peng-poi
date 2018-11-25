package com.github.zzlhy.main;

import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.func.GeneratorDataHandler;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

/**
 * Excel导出接口
 * Created by zpeng on 2018-11-15.
 */
public interface ExportAbstract {


    /**
     * 普通对象方式导出 （适用数据量较少情况）
     * @param tableParam Excel参数对象
     * @param data data
     * @param exportStyle 单元格样式
     * @return Workbook
     */
    Workbook exportByObject(TableParam tableParam, List<?> data,ExportStyle exportStyle);

    /**
     * 大量数据导出 (说明:支持百万数据导出,支持多sheet写入,分页获取数据,减少内存消耗)
     * @param tableParam Excel参数对象
     * @param total 数据总条数
     * @param generatorDataHandler 生成数据的方法
     * @param exportStyle 单元格样式
     * @return SXSSFWorkbook
     */
    SXSSFWorkbook exportBigData(TableParam tableParam, long total, GeneratorDataHandler generatorDataHandler,ExportStyle exportStyle);

}
