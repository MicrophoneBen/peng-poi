package com.github.zzlhy;

import com.github.zzlhy.entity.ExcelType;
import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.func.GeneratorDataHandler;
import com.github.zzlhy.func.IndexChangeHandler;
import com.github.zzlhy.func.SaveDataHandler;
import com.github.zzlhy.func.ValidateDataHandler;
import com.github.zzlhy.main.ExcelExport;
import com.github.zzlhy.main.ExcelImport;
import com.github.zzlhy.main.ExportStyle;
import com.github.zzlhy.main.ExportStyleImpl;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.beans.IntrospectionException;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.ParseException;
import java.util.List;
import java.util.Map;

/**
 *
 * Created by Administrator on 2018-10-19.
 */
public class ExcelUtils {


    /**
     * 大量数据导出 (分页查询写入)
     * @param tableParam 表格参数
     * @param total 总页数
     * @param generatorDataHandler 生成数据的方法,需要自己实现
     * @return 导出完成后需要调用SXSSFWorkbook的dispose()方法删除临时文件
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static SXSSFWorkbook excelExport(TableParam tableParam, long total, GeneratorDataHandler generatorDataHandler) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        return ExcelExport.exportExcelBigData(tableParam,total,generatorDataHandler,new ExportStyleImpl());
    }

    /**
     * 大量数据导出 (分页查询写入)
     * @param tableParam 表格参数
     * @param total 总页数
     * @param generatorDataHandler 生成数据的方法,需要自己实现
     * @param exportStyle 单元格样式实现
     * @return 导出完成后需要调用SXSSFWorkbook的dispose()方法删除临时文件
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static SXSSFWorkbook excelExport(TableParam tableParam, long total, GeneratorDataHandler generatorDataHandler, ExportStyle exportStyle) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        return ExcelExport.exportExcelBigData(tableParam,total,generatorDataHandler,exportStyle);
    }

    /**
     * 大量数据导出(基于对象,一次写出全部数据,数据量过大时,会占用大量内存)
     * @param tableParam 表格参数
     * @param data 数据
     * @return SXSSFWorkbook
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static SXSSFWorkbook excelExport(TableParam tableParam, List<?> data) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        tableParam.setExcelType(ExcelType.SXLSX);
        return (SXSSFWorkbook) ExcelExport.exportExcelByObject(tableParam,data,new ExportStyleImpl());
    }

    /**
     * 大量数据导出(基于对象,一次写出全部数据,数据量过大时,会占用大量内存)
     * @param tableParam 表格参数
     * @param data 数据
     * @param exportStyle 单元格样式实现
     * @return SXSSFWorkbook
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static SXSSFWorkbook excelExport(TableParam tableParam, List<?> data, ExportStyle exportStyle) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        tableParam.setExcelType(ExcelType.SXLSX);
        return (SXSSFWorkbook) ExcelExport.exportExcelByObject(tableParam,data,exportStyle);
    }

    /**
     * 普通导出(基于对象)
     * @param tableParam tableParam
     * @param data data
     * @return Workbook
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static Workbook excelExportByObject(TableParam tableParam, List<?> data) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        return ExcelExport.exportExcelByObject(tableParam,data,new ExportStyleImpl());
    }

    /**
     * 普通导出(基于对象)
     * @param tableParam tableParam
     * @param data data
     * @param exportStyle 单元格样式实现
     * @return Workbook
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static Workbook excelExportByObject(TableParam tableParam, List<?> data, ExportStyle exportStyle) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        return ExcelExport.exportExcelByObject(tableParam,data,exportStyle);
    }

    /**
     * 普通导出(基于List-Map)
     * @param tableParam tableParam
     * @param data data
     * @return Workbook
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static Workbook excelExportByMap(TableParam tableParam, List<? extends Map<?,?>> data) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        return ExcelExport.exportExcelByMap(tableParam,data,new ExportStyleImpl());
    }

    /**
     * 普通导出(基于List-Map)
     * @param tableParam tableParam
     * @param data data
     * @param exportStyle 单元格样式实现
     * @return Workbook
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     * @throws InvocationTargetException e
     */
    public static Workbook excelExportByMap(TableParam tableParam, List<? extends Map<?,?>> data, ExportStyle exportStyle) throws IllegalAccessException, IntrospectionException, InvocationTargetException {
        return ExcelExport.exportExcelByMap(tableParam,data,exportStyle);
    }

    /**
     * Excel读取 （一次返回所有读取的数据）
     * @param stream stream
     * @param tableParam tableParam
     * @param clazz clazz
     * @return 读取的数据
     * @throws IllegalAccessException e
     * @throws ParseException e
     * @throws IntrospectionException e
     * @throws IOException e
     * @throws InstantiationException e
     * @throws InvalidFormatException e
     * @throws InvocationTargetException e
     */
    public static List<?> readExcel(InputStream stream,TableParam tableParam, Class<?> clazz) throws IllegalAccessException, ParseException, IntrospectionException, IOException, InstantiationException, InvalidFormatException, InvocationTargetException {
        return ExcelImport.readExcel(stream,tableParam,clazz);
    }

    /**
     * Excel读取 （一次返回所有读取的数据）
     * @param filePath Excel 文件路劲
     * @param tableParam 表格参数
     * @param clazz 实体类
     * @return 读取的数据
     * @throws IllegalAccessException e
     * @throws ParseException e
     * @throws IntrospectionException e
     * @throws IOException e
     * @throws InstantiationException e
     * @throws InvalidFormatException e
     * @throws InvocationTargetException e
     */
    public static List<?> readExcel(String filePath,TableParam tableParam, Class<?> clazz) throws IllegalAccessException, ParseException, IntrospectionException, IOException, InstantiationException, InvalidFormatException, InvocationTargetException {
        FileInputStream stream=new FileInputStream(filePath);
        return ExcelImport.readExcel(stream,tableParam,clazz);
    }


    /**
     * Excel导入 （最全面的的导入）
     * @param stream 流
     * @param tableParam 表格参数
     * @param clazz 实体类
     * @param saveDataHandler 保存数据方法
     * @param validateDataHandler 验证数据方法
     * @param indexChangeHandler 执行记录数变化时的回调方法
     * @return 导入失败的数据
     * @throws IllegalAccessException e
     * @throws ParseException e
     * @throws IntrospectionException e
     * @throws IOException e
     * @throws InstantiationException e
     * @throws InvalidFormatException e
     * @throws InvocationTargetException e
     */
    public static List<Map<String,Object>> excelImport(InputStream stream, TableParam tableParam, Class<?> clazz, SaveDataHandler saveDataHandler, ValidateDataHandler validateDataHandler, IndexChangeHandler indexChangeHandler) throws IllegalAccessException, ParseException, IntrospectionException, IOException, InstantiationException, InvalidFormatException, InvocationTargetException {
        return ExcelImport.importExcel(stream,tableParam,clazz,saveDataHandler,validateDataHandler,indexChangeHandler);
    }

    /**
     * Excel导入 (无执行记录变化回调)
     * @param stream 流
     * @param tableParam 表格参数
     * @param clazz 实体类
     * @param saveDataHandler 保存数据方法
     * @param validateDataHandler 验证数据方法
     * @return 导入失败的数据
     * @throws IllegalAccessException e
     * @throws ParseException e
     * @throws IntrospectionException e
     * @throws IOException e
     * @throws InstantiationException e
     * @throws InvalidFormatException e
     * @throws InvocationTargetException e
     */
    public static List<?> excelImport(InputStream stream, TableParam tableParam, Class<?> clazz, SaveDataHandler saveDataHandler, ValidateDataHandler validateDataHandler) throws IllegalAccessException, ParseException, IntrospectionException, IOException, InstantiationException, InvalidFormatException, InvocationTargetException {
        return ExcelImport.importExcel(stream,tableParam,clazz,saveDataHandler,validateDataHandler,null);
    }

    /**
     * Excel导入 (无验证和执行记录变化回调)
     * @param stream 流
     * @param tableParam 表格参数
     * @param clazz 实体类
     * @param saveDataHandler 保存数据方法
     * @return 导入失败的数据
     * @throws IllegalAccessException e
     * @throws ParseException e
     * @throws IntrospectionException e
     * @throws IOException e
     * @throws InstantiationException e
     * @throws InvalidFormatException e
     * @throws InvocationTargetException e
     */
    public static List<?> excelImport(InputStream stream, TableParam tableParam, Class<?> clazz, SaveDataHandler saveDataHandler) throws IllegalAccessException, ParseException, IntrospectionException, IOException, InstantiationException, InvalidFormatException, InvocationTargetException {
        return ExcelImport.importExcel(stream,tableParam,clazz,saveDataHandler,null,null);
    }


}
