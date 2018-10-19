package com.github.zzlhy.main;


import com.github.zzlhy.entity.Col;
import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.func.ConvertValue;
import com.github.zzlhy.func.IndexChangeHandler;
import com.github.zzlhy.func.SaveDataHandler;
import com.github.zzlhy.func.ValidateDataHandler;
import com.github.zzlhy.util.Utils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Excel 导入
 * Created by Administrator on 2018-10-15.
 */
public class ExcelImport {


    //默认日期转换格式
    private static final String FORMAT="yyyy-MM-dd HH:mm:ss";
    private static SimpleDateFormat sdf= new SimpleDateFormat(FORMAT);

    /**
     * 导入-- 支持xls和xlsx  (说明：支持大量数据导入,减少内存消耗,支持多个sheet读取)
     * @param stream 文件流
     * @param tableParam Excel配置对象
     * @param clazz Class
     * @param saveDataHandler 保存数据的方法
     * @param validateDataHandler 效验数据的方法
     * @param indexChangeHandler 处理条数变化的回调
     * @return 导入失败的数据
     */
    public static List<Map<String,Object>> importExcel(InputStream stream, TableParam tableParam, Class<?> clazz, SaveDataHandler saveDataHandler, ValidateDataHandler validateDataHandler, IndexChangeHandler indexChangeHandler) throws IOException, IllegalAccessException, InstantiationException, IntrospectionException, InvocationTargetException, ParseException, InvalidFormatException {
        //数据开始行
        int startRow=tableParam.getReadRow();
        //列参数配置
        List<Col> cols = tableParam.getCols();
        //失败数据
        List<Map<String,Object>> failData=new ArrayList();
        //已处理的条数
        int index = 0;

        Workbook workbook= WorkbookFactory.create(stream);
        //遍历Sheet
        for (Sheet sheet : workbook) {
            for(int i = startRow;i <= sheet.getLastRowNum();i++){
                Row row = sheet.getRow(i);
                if(row == null){
                    break;
                }
                //创建对象实例
                Object instance = clazz.newInstance();
                Map<String, Object> objectMap = new HashMap<String, Object>();
                //循环赋值
                for (int j = 0;j < cols.size();j++) {
                    //获取属性
                    PropertyDescriptor propertyDescriptor = new PropertyDescriptor(cols.get(j).getKey(), clazz);
                    //获取set方法
                    Method writeMethod = propertyDescriptor.getWriteMethod();
                    //对象属性类型
                    Class<?> propertyType = propertyDescriptor.getPropertyType();

                    Cell cell = row.getCell(j);

                    //列的值
                    String value = getCellValue(cell);


                    //判断是否需要转换--传了此参数说明需要转换
                    ConvertValue convertValue = cols.get(j).getConvertValue();
                    if(convertValue!=null){
                        value = convertValue.convert(value);
                    }

                    //对象属性赋值
                    setValue(writeMethod,propertyType,instance,cols.get(j).getFormat(),value);
                    objectMap.put(cols.get(j).getKey(),value);
                }

                //验证数据并保存数据
                if(validateDataHandler != null) {
                    String desc = validateDataHandler.valid(instance);
                    if (Utils.notEmpty(desc)) {
                        //验证未通过的处理
                        objectMap.put("errorMsg",desc);
                        failData.add(objectMap);
                    } else {
                        //保存数据
                        try {
                            saveDataHandler.save(instance);
                        }catch (Exception e){
                            objectMap.put("errorMsg",e.getMessage());
                            failData.add(objectMap);
                        }
                    }
                }else{
                    //保存数据
                    try {
                        saveDataHandler.save(instance);
                    }catch (Exception e){
                        objectMap.put("errorMsg",e.getMessage());
                        failData.add(objectMap);
                    }
                }

                index++;
                //处理条数变化回调
                if(indexChangeHandler != null) {
                    indexChangeHandler.change(index);
                }
            }
        }

        return failData;
    }

    /**
     * 导入-- 支持xls和xlsx  (说明：支持大量数据导入,减少内存消耗,支持多个sheet读取)  不需要处理条数变化回调
     * @param stream 文件流
     * @param tableParam Excel配置对象
     * @param clazz Class
     * @param saveDataHandler 保存数据的方法
     * @param validateDataHandler 效验数据的方法
     * @return 导入失败的数据
     */
    public static List<?> importExcel(InputStream stream, TableParam tableParam, Class<?> clazz, SaveDataHandler saveDataHandler, ValidateDataHandler validateDataHandler) throws IOException, IllegalAccessException, InstantiationException, IntrospectionException, InvocationTargetException, ParseException, InvalidFormatException {
        return importExcel(stream,tableParam,clazz,saveDataHandler,validateDataHandler,null);
    }

    /**
     * 普通导入 -- 支持xls和xlsx
     * @param stream 文件流
     * @param tableParam Excel配置对象
     * @param clazz Class
     * @return 返回读取到的数据列表
     */
    public static List<?> readExcel(InputStream stream,TableParam tableParam, Class<?> clazz) throws IOException, IllegalAccessException, InstantiationException, IntrospectionException, InvocationTargetException, ParseException, InvalidFormatException {
        List<Object> list=new ArrayList();
        Workbook workbook=WorkbookFactory.create(stream);
        //获取第一个sheet
        Sheet sheet= workbook.getSheetAt(0);

        //数据开始行
        int startRow=tableParam.getReadRow();

        for(int i=startRow;i<=sheet.getLastRowNum();i++){
            Row row = sheet.getRow(i);
            if(row == null){
                break;
            }
            Object instance = clazz.newInstance();
            List<Col> cols = tableParam.getCols();
            for (int j=0;j<cols.size();j++) {
                //获取属性
                PropertyDescriptor propertyDescriptor = new PropertyDescriptor(cols.get(j).getKey(), clazz);
                //获取set方法
                Method writeMethod = propertyDescriptor.getWriteMethod();
                Cell cell = row.getCell(j);
                Class<?> propertyType = propertyDescriptor.getPropertyType();

                //列的值
                String value = getCellValue(cell);

                //判断是否需要转换--传了此参数说明需要转换
                ConvertValue convertValue = cols.get(j).getConvertValue();
                if(convertValue!=null){
                    value = convertValue.convert(value);
                }

                //对象属性赋值
                setValue(writeMethod,propertyType,instance,cols.get(j).getFormat(),value);
            }

            list.add(instance);
        }
        return list;
    }

    /**
     * 获取列的值
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell){
        String value = "";
        if(cell!=null) {
            //判断格式
            switch (cell.getCellType()) {
                case STRING:
                    value = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    value = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    //有公式时的处理
                    cell.setCellType(CellType.STRING);
                    value = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    //判断是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        value = sdf.format(cell.getDateCellValue());
                    } else {
                        // 不是日期格式，则防止当数字过长时以科学计数法显示
                        cell.setCellType(CellType.STRING);
                        value = cell.toString();
                    }
                    break;
                default:
                    value = cell.getStringCellValue();
                    break;
            }
        }
        return value;
    }

    /**
     * 对象赋值
     * @param writeMethod  属性set方法
     * @param propertyType 属性类型
     * @param instance 对象
     * @param format 日期格式化字符串
     * @param value 值
     */
    private static void setValue(Method writeMethod,Class<?> propertyType,Object instance,String format,String value) throws InvocationTargetException, IllegalAccessException, ParseException {
        //判断类型---说明：如果传进来为空或空字符串时,统一把值设置为(Object)null,否则invoke赋值会报错
        if(Utils.isEmpty(value)){
            writeMethod.invoke(instance, (Object)null);
        }else {
            if (propertyType == String.class) {
                writeMethod.invoke(instance, value);
            } else if (propertyType == Integer.class) {
                writeMethod.invoke(instance, Integer.parseInt(value));
            } else if (propertyType == Short.class) {
                writeMethod.invoke(instance, Short.parseShort(value));
            } else if (propertyType == Boolean.class) {
                writeMethod.invoke(instance, Boolean.parseBoolean(value));
            } else if (propertyType == Long.class) {
                writeMethod.invoke(instance, Long.parseLong(value));
            } else if (propertyType == Double.class) {
                writeMethod.invoke(instance, Double.parseDouble(value));
            } else if (propertyType == Float.class) {
                writeMethod.invoke(instance, Float.parseFloat(value));
            } else if (propertyType == BigDecimal.class) {
                writeMethod.invoke(instance, BigDecimal.valueOf(Double.parseDouble(value)));
            }else if (propertyType == Date.class) {
                if (Utils.isEmpty(format)) {
                    sdf = new SimpleDateFormat(format);
                }
                Date date = sdf.parse(value);
                writeMethod.invoke(instance, date);
            }else if (propertyType == LocalDate.class) {
                if (Utils.notEmpty(format)) {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
                    LocalDate localDate = LocalDate.parse(value, formatter);
                    writeMethod.invoke(instance, localDate);
                }else{
                    writeMethod.invoke(instance, LocalDate.parse(value, DateTimeFormatter.ofPattern("yyyy-MM-dd")));
                }

            }else if (propertyType == LocalDateTime.class) {
                if (Utils.notEmpty(format)) {
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(format);
                    LocalDateTime localDateTime = LocalDateTime.parse(value, formatter);
                    writeMethod.invoke(instance, localDateTime);
                }else{
                    writeMethod.invoke(instance, LocalDateTime.parse(value, DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")));
                }
            }
        }
    }


}
