package com.github.zzlhy.main;


import com.github.zzlhy.entity.Col;
import com.github.zzlhy.entity.ExcelType;
import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.func.ConvertValue;
import com.github.zzlhy.func.GeneratorDataHandler;
import com.github.zzlhy.util.Utils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Excel 导出  (支持大量数据导出)
 * Created by Administrator on 2018-10-11.
 */
public class ExcelExport {

    /**
     * 数据导出 -- 对象方式 (说明:普通导出,数据量较少情况导出)
     * @param tableParam Excel参数对象
     * @param data data
     * @return Workbook
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    public static Workbook exportExcelByObject(TableParam tableParam, List<?> data) throws InvocationTargetException, IllegalAccessException, IntrospectionException {

        /*创建Workbook和Sheet*/
        Workbook workbook;
        if(ExcelType.XLSX.equals(tableParam.getExcelType())){
            workbook = new XSSFWorkbook();
        }else if(ExcelType.XLS.equals(tableParam.getExcelType())){
            workbook = new HSSFWorkbook();
        }else {
            workbook = new SXSSFWorkbook(100);
        }

        /*创建Workbook和Sheet*/
        Sheet sheet = workbook.createSheet(tableParam.getSheetName());//创建工作表(Sheet)
        //开始行
        Integer startRow = tableParam.getStartRow();

        //创建标题
        if(tableParam.getCreateHeadRow()){
            /*创建标题行*/
            Row row = sheet.createRow(startRow);// 创建行,从0开始
            //标题设置
            setHeadRow(workbook,row,tableParam);
        }
        //当前数据处理到的行数,开始时为标题行的下一行（每个sheet重新计算）
        int currentRow = startRow+1;
        addRows(sheet, currentRow, tableParam, data);
        return workbook;
    }

    /**
     * 数据导出 -- Map方式 (说明:普通导出,数据量较少情况导出)
     * @param tableParam Excel参数对象
     * @param data data
     * @return Workbook
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    public static Workbook exportExcelByMap(TableParam tableParam, List<? extends Map<?,?>> data) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
        /*创建Workbook和Sheet*/
        Workbook workbook;
        if(ExcelType.XLSX.equals(tableParam.getExcelType())){
            workbook = new XSSFWorkbook();
        }else if(ExcelType.XLS.equals(tableParam.getExcelType())){
            workbook = new HSSFWorkbook();
        }else {
            workbook = new SXSSFWorkbook(100);
        }

        /*创建Workbook和Sheet*/
        Sheet sheet = workbook.createSheet(tableParam.getSheetName());//创建工作表(Sheet)

        //开始行
        Integer startRow = tableParam.getStartRow();

        //创建标题
        if(tableParam.getCreateHeadRow()){
            /*创建标题行*/
            Row row = sheet.createRow(startRow);// 创建行,从0开始
            //标题设置
            setHeadRow(workbook,row,tableParam);
        }
        //当前数据处理到的行数,开始时为标题行的下一行（每个sheet重新计算）
        int currentRow = startRow+1;
        addRowsByMap(sheet, currentRow, tableParam, data);
        return workbook;
    }

    /**
     * 大量数据导出 -- 对象方式 (说明:支持百万数据导出,支持多sheet写入，减少内存消耗)
     * @param tableParam Excel参数对象
     * @param total 数据总条数
     * @param generatorDataHandler 生成数据的方法
     * @return SXSSFWorkbook
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    public static SXSSFWorkbook exportExcelBigData(TableParam tableParam, long total,GeneratorDataHandler generatorDataHandler) throws InvocationTargetException, IllegalAccessException, IntrospectionException {

        /*创建Workbook*/
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);//在内存中保持100行，超过100行将被刷新到磁盘

        //每个sheet允许写入的数据条数
        int sheetDataTotal = tableParam.getSheetDataTotal();
        //sheet个数
        long sheetCount = 1;
        //判断数据总行数,决定是否创建多个sheet  每超过10W条则新加一个sheet
        if(total > sheetDataTotal){
            //计算需要创建sheet的个数
            sheetCount = total/sheetDataTotal;
            if(total%sheetDataTotal != 0L){
                sheetCount+=1;
            }
        }

        //分页查询时的每页条数,固定1000条,外部查询时也需要每次为1000条的返回
        int pageSize = 1000;
        //总页数计算
        long page = total / pageSize;
        //当前页码
        int currentPage = 0;

        //循环创建sheet
        for (int j = 0; j < sheetCount; j++) {

            //创建工作表(Sheet)
            Sheet sheet = workbook.createSheet(tableParam.getSheetName()+j);

            //开始写入的行号
            Integer startRow = tableParam.getStartRow();

            //创建标题行
            if(tableParam.getCreateHeadRow()){
                    /*创建标题行*/
                Row row = sheet.createRow(startRow);// 创建行,从0开始
                //标题设置
                setHeadRow(workbook,row,tableParam);
            }

            //当前数据处理到的行数,开始时为标题行的下一行（每个sheet重新计算）
            int currentRow = startRow+1;

            //本次执行了多少页
            int index = 0;
            List<?> list;
            //开始循环加入数据(分页查询数据)
            for (int i = currentPage; i <= page; i++) {
                list = generatorDataHandler.generatorData(i,pageSize);
                int tempRow = addRows(sheet, currentRow, tableParam, list);
                list.clear();
                currentRow = tempRow;
                //如果超过了限制的行数则跳出,进行新的sheet写入
                if(currentRow >= sheetDataTotal ){
                    break;
                }
                index++;
            }
            currentPage+=index;
        }

        return workbook;
    }

    /**
     * 新增数据行  说明：传入的是对象集合
     * @param sheet sheet
     * @param currentRow  当前处理到的行号
     * @param tableParam  配置参数
     * @param data        数据
     * @return 当前处理到的行号
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    private static int addRows(Sheet sheet, int currentRow, TableParam tableParam, List<?> data) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
        //创建数据
        for(int k = 0; k<data.size(); k++) {
            // 创建行
            Row _row = sheet.createRow(currentRow);
            //设置行的高度
            _row.setHeightInPoints(tableParam.getHeight());

            for (int j=0;j<tableParam.getCols().size();j++) {
                //设置列宽
                sheet.setColumnWidth(j,tableParam.getCols().get(j).getWidth()*256);

                // 创建行的单元格
                Cell cell = _row.createCell(j);

                //判断列是否为设置公式的列,设置公式不根据key生成数据,直接写入公式即可
                if(Utils.notEmpty(tableParam.getCols().get(j).getFormula())){
                    cell.setCellFormula(tableParam.getCols().get(j).getFormula());
                    continue;
                }


                //获取属性
                PropertyDescriptor propertyDescriptor = new PropertyDescriptor(tableParam.getCols().get(j).getKey(), data.get(k).getClass());
                //获取get方法
                Method readMethod = propertyDescriptor.getReadMethod();
                //字段类型
                //Class<?> propertyType = propertyDescriptor.getPropertyType();
                //执行
                Object result = readMethod.invoke(data.get(k));

                String format = tableParam.getCols().get(j).getFormat();//获取日期的格式化的格式
                ConvertValue convertValue = tableParam.getCols().get(j).getConvertValue();//需要转换值的方法对象
                setCell(cell,result,format,convertValue);

            }
            currentRow++;
        }
        return currentRow;
    }

    /**
     * 新增数据行  说明：传入的是Map集合
     * @param sheet sheet
     * @param currentRow  当前处理到的行号(从哪行开始创建行)
     * @param tableParam  配置参数
     * @param data        数据
     * @return 当前处理到的行号
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    private static int addRowsByMap(Sheet sheet, int currentRow, TableParam tableParam, List<? extends Map<?,?>> data) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
        //创建数据
        for(int k = 0; k<data.size(); k++) {
            // 创建行
            Row _row = sheet.createRow(currentRow);
            //设置行的高度
            _row.setHeightInPoints(tableParam.getHeight());

            for (int j=0;j<tableParam.getCols().size();j++) {
                //设置列宽
                sheet.setColumnWidth(j,tableParam.getCols().get(j).getWidth()*256);

                // 创建行的单元格
                Cell cell = _row.createCell(j);

                //判断列是否为设置公式的列,设置公式不根据key生成数据,直接写入公式即可
                if(Utils.notEmpty(tableParam.getCols().get(j).getFormula())){
                    cell.setCellFormula(tableParam.getCols().get(j).getFormula());
                    continue;
                }

                //根据Map的key获取value值
                Object result = data.get(k).get(tableParam.getCols().get(j).getKey());

                String format = tableParam.getCols().get(j).getFormat();//获取日期的格式化的格式
                ConvertValue convertValue = tableParam.getCols().get(j).getConvertValue();//需要转换值的方法对象
                setCell(cell,result,format,convertValue);

            }
            currentRow++;
        }
        return currentRow;
    }

    /**
     * 标题行设置
     * @param workbook workbook
     * @param row row
     * @param tableParam tableParam
     */
    private static void setHeadRow(Workbook workbook, Row row, TableParam tableParam){
        //标题行样式
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        //加粗
        font.setBold(tableParam.getHeadRowStyle().getHeadBold());
        style.setFont(font);
        //居中
        style.setAlignment(tableParam.getHeadRowStyle().getHorizontalAlignment());
        List<Col> columnParams = tableParam.getCols();
        for(int i=0;i<columnParams.size();i++){
            // 创建行的单元格,也是从0开始
            Cell cell = row.createCell(i);
            // 设置单元格内容
            cell.setCellValue(columnParams.get(i).getTitle());
            cell.setCellStyle(style);

            //数据有效性设置(数据格式设置)
            //int lastRow = 65535;
            //if(tableParam.getExcelType().equals(ExcelType.XLSX) || tableParam.getExcelType().equals(ExcelType.SXLSX)){
            //    lastRow = 1048576;
            //}
            //CellRangeAddressList regions = new CellRangeAddressList(row.getRowNum(), lastRow, i, i);//选定一个区域  参数：起始行序号，终止行序号，起始列序号，终止列序号
            ////DVConstraint constraint = DVConstraint.createDateConstraint(DVConstraint.OperatorType.BETWEEN, "1993-01-01", "2014-12-31", "yyyy-MM-dd");
            ////下拉菜单生成
            //DVConstraint constraint = DVConstraint.createExplicitListConstraint(new String[] { "C++", "Java", "C#" });
            //HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            //Sheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
            //sheet.addValidationData(dataValidate);

        }
    }

    //默认日期转换格式
    private static final String FORMAT="yyyy-MM-dd";
    private static final String FORMAT2="yyyy-MM-dd HH:mm:ss";
    private static SimpleDateFormat sdfDate= new SimpleDateFormat(FORMAT);
    private static SimpleDateFormat sdfTime= new SimpleDateFormat(FORMAT2);

    /**
     * 单元格cell值设置
     * @param cell cell
     * @param result result
     */
    private static void setCell(Cell cell, Object result, String format, ConvertValue convertValue){
        //设置单元格值及属性
        String resultStr = String.valueOf(result);
        if(result instanceof String){
            if(result != null){
                cell.setCellValue(resultStr);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Integer){
            if(result != null) {
                cell.setCellValue(Integer.parseInt(resultStr));
            }else{
                Integer empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Double){
            if(result != null) {
                cell.setCellValue(Double.parseDouble(resultStr));
            }else{
                Double empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Short){
            if(result != null) {
                //如果需要值替换
                if (convertValue != null) {
                    String val = convertValue.convert(result);
                    cell.setCellValue(val);
                } else {
                    cell.setCellValue(Short.parseShort(resultStr));
                }
            }else{
                Short empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Long){
            if(result != null) {
                cell.setCellValue(Long.parseLong(resultStr));
            }else{
                Long empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Float){
            if(result != null) {
                cell.setCellValue(Float.parseFloat(resultStr));
            }else{
                Float empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof BigDecimal){
            if(result != null) {
                cell.setCellValue(Double.parseDouble(resultStr));
            }else{
                Double empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Boolean){
            if(result != null) {
                //如果需要值替换
                if (convertValue != null) {
                    String val = convertValue.convert(result);
                    cell.setCellValue(val);
                } else {
                    cell.setCellValue(Boolean.parseBoolean(resultStr));
                }
            }else{
                Boolean empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof Date){
            if(result != null) {
                //日期直接用字符串输出
                Date date = (Date) result;
                String strDate;
                if (Utils.notEmpty(format)) {
                    SimpleDateFormat s = new SimpleDateFormat(format);
                    strDate = s.format(date);
                }else{
                    strDate = sdfTime.format(date);
                }
                cell.setCellValue(strDate);
            }else{
                Date empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof LocalDate){
            if(result != null) {
                LocalDate localDate = (LocalDate) result;
                //LocalDate转换为Date
                ZoneId zoneId = ZoneId.systemDefault();
                ZonedDateTime zdt = localDate.atStartOfDay(zoneId);
                Date date = Date.from(zdt.toInstant());
                String strDate = sdfDate.format(date);
                cell.setCellValue(strDate);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
        }else if(result instanceof LocalDateTime){
            if(result != null) {
                LocalDateTime localDateTime = (LocalDateTime) result;
                //LocalDateTime转换为Date
                ZoneId zoneId = ZoneId.systemDefault();
                ZonedDateTime zdt =localDateTime.atZone(zoneId);
                Date date = Date.from(zdt.toInstant());
                String strDate = sdfTime.format(date);
                cell.setCellValue(strDate);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
        }else{
            if(result != null) {
                cell.setCellValue(resultStr);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
        }
    }

}
