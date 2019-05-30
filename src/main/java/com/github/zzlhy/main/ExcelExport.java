package com.github.zzlhy.main;


import com.github.zzlhy.entity.Col;
import com.github.zzlhy.entity.DropdownParam;
import com.github.zzlhy.entity.ExcelType;
import com.github.zzlhy.entity.TableParam;
import com.github.zzlhy.func.ConvertValue;
import com.github.zzlhy.func.GeneratorDataHandler;
import com.github.zzlhy.util.Utils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
import java.util.ArrayList;
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
     * @param exportStyle 单元格样式实现
     * @return Workbook
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    public static Workbook exportExcelByObject(TableParam tableParam, List<?> data,ExportStyle exportStyle) throws InvocationTargetException, IllegalAccessException, IntrospectionException {

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
        //冻结列
        sheet.createFreezePane(tableParam.getFreezeColSplit(),tableParam.getFreezeRowSplit());
        //合并单元格
        if(tableParam.getMergeRegion() != null && tableParam.getMergeRegion().size() > 0){
            for (CellRangeAddress cellAddresses : tableParam.getMergeRegion()) {
                sheet.addMergedRegion(cellAddresses);
            }
        }

        //开始行
        Integer writeRow = tableParam.getWriteRow();

        //标题行设置、下拉列表设置
        rowConfig(workbook,sheet,tableParam);

        //样式对象创建,可复用
        for (int j=0;j<tableParam.getCols().size();j++) {
            //样式获取设置
            if(exportStyle != null){
                CellStyle cellStyle = exportStyle.createCellStyle(workbook, tableParam.getCols().get(j).getColStyle());
                if(cellStyle != null) {
                    tableParam.getCols().get(j).setCellStyle(cellStyle);
                    //日期格式设置
                    DataFormat dataFormat = workbook.createDataFormat();
                    tableParam.getCols().get(j).setDataFormat(dataFormat);
                }
            }
        }

        //当前数据处理到的行数,开始时为标题行的下一行（每个sheet重新计算）
        int currentRow = writeRow+1;
        addRows(workbook, sheet, currentRow, tableParam, exportStyle, data);
        return workbook;
    }

    /**
     * 数据导出 -- Map方式 (说明:普通导出,数据量较少情况导出)
     * @param tableParam Excel参数对象
     * @param data data
     * @param exportStyle 单元格样式实现
     * @return Workbook
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    public static Workbook exportExcelByMap(TableParam tableParam, List<? extends Map<?,?>> data,ExportStyle exportStyle) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
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
        //冻结列
        sheet.createFreezePane(tableParam.getFreezeColSplit(),tableParam.getFreezeRowSplit());
        //合并单元格
        if(tableParam.getMergeRegion() != null && tableParam.getMergeRegion().size() > 0){
            for (CellRangeAddress cellAddresses : tableParam.getMergeRegion()) {
                sheet.addMergedRegion(cellAddresses);
            }
        }

        //开始行
        Integer writeRow = tableParam.getWriteRow();

        //标题行设置、下拉列表设置
        rowConfig(workbook,sheet,tableParam);

        //样式对象创建,可复用
        for (int j=0;j<tableParam.getCols().size();j++) {
            //样式获取设置
            if(exportStyle != null){
                CellStyle cellStyle = exportStyle.createCellStyle(workbook, tableParam.getCols().get(j).getColStyle());
                if(cellStyle != null) {
                    tableParam.getCols().get(j).setCellStyle(cellStyle);
                    //日期格式设置
                    DataFormat dataFormat = workbook.createDataFormat();
                    tableParam.getCols().get(j).setDataFormat(dataFormat);
                }
            }
        }

        //当前数据处理到的行数,开始时为标题行的下一行（每个sheet重新计算）
        int currentRow = writeRow+1;
        addRowsByMap(workbook, sheet, currentRow, tableParam, exportStyle, data);
        return workbook;
    }

    /**
     * 大量数据导出 -- 对象方式 (说明:支持百万数据导出,支持多sheet写入，减少内存消耗)
     * @param tableParam Excel参数对象
     * @param total 数据总条数
     * @param generatorDataHandler 生成数据的方法
     * @param exportStyle 单元格样式实现
     * @return SXSSFWorkbook
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    public static SXSSFWorkbook exportExcelBigData(TableParam tableParam, long total,GeneratorDataHandler generatorDataHandler, ExportStyle exportStyle) throws InvocationTargetException, IllegalAccessException, IntrospectionException {

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

        //分页查询时的每页条数,固定10000条,外部查询时也需要每次为10000条的返回
        int pageSize = 10000;
        //总页数计算
        long page;
        if(total % pageSize == 0){
            page = total / pageSize;
        }else{
            page = total / pageSize + 1;
        }
        //当前页码
        int currentPage = 0;

        //样式对象创建,可复用
        for (int j=0;j<tableParam.getCols().size();j++) {
            if(tableParam.getCols().get(j).getCellStyle() == null){
                //单元格样式
                CellStyle cellStyle = workbook.createCellStyle();
                tableParam.getCols().get(j).setCellStyle(cellStyle);
            }

            //样式获取设置
            if(exportStyle != null){
                CellStyle cellStyle = exportStyle.createCellStyle(workbook, tableParam.getCols().get(j).getColStyle());
                if(cellStyle != null) {
                    tableParam.getCols().get(j).setCellStyle(cellStyle);
                }
            }
            //日期格式设置
            DataFormat dataFormat = workbook.createDataFormat();
            tableParam.getCols().get(j).setDataFormat(dataFormat);
        }


        //循环创建sheet
        for (int j = 0; j < sheetCount; j++) {

            //创建工作表(Sheet)
            Sheet sheet = workbook.createSheet(tableParam.getSheetName()+j);
            //冻结列
            sheet.createFreezePane(tableParam.getFreezeColSplit(),tableParam.getFreezeRowSplit());
            //合并单元格
            if(tableParam.getMergeRegion() != null && tableParam.getMergeRegion().size() > 0){
                for (CellRangeAddress cellAddresses : tableParam.getMergeRegion()) {
                    sheet.addMergedRegion(cellAddresses);
                }
            }

            //开始写入的行号
            Integer writeRow = tableParam.getWriteRow();

            //标题行设置
            rowConfig(workbook,sheet,tableParam);

            //当前数据处理到的行数,开始时为标题行的下一行（每个sheet重新计算）
            int currentRow = writeRow+1;

            //本次执行了多少页
            int index = 0;
            List<?> list;
            //开始循环加入数据(分页查询数据)
            for (int i = currentPage; i <= page; i++) {
                list = generatorDataHandler.generatorData(i,pageSize);
                int tempRow = addRows(workbook, sheet, currentRow, tableParam,exportStyle, list);
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
     * @param workbook workbook
     * @param sheet sheet
     * @param currentRow  当前处理到的行号
     * @param tableParam  配置参数
     * @param exportStyle 单元格样式实现
     * @param data        数据
     * @return 当前处理到的行号
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    private static int addRows(Workbook workbook,Sheet sheet, int currentRow, TableParam tableParam, ExportStyle exportStyle, List<?> data) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
        //创建数据
        for(int k = 0; k<data.size(); k++) {
            // 创建行
            Row _row = sheet.createRow(currentRow);
            //设置行的高度
            _row.setHeightInPoints(tableParam.getHeight());

            for (int j=0;j<tableParam.getCols().size();j++) {
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

                setCell(cell,result,tableParam.getCols().get(j));
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
     * @param exportStyle 单元格样式实现
     * @param data        数据
     * @return 当前处理到的行号
     * @throws InvocationTargetException e
     * @throws IllegalAccessException e
     * @throws IntrospectionException e
     */
    private static int addRowsByMap(Workbook workbook, Sheet sheet, int currentRow, TableParam tableParam, ExportStyle exportStyle, List<? extends Map<?,?>> data) throws InvocationTargetException, IllegalAccessException, IntrospectionException {
        //创建数据
        for(int k = 0; k<data.size(); k++) {
            // 创建行
            Row _row = sheet.createRow(currentRow);
            //设置行的高度
            _row.setHeightInPoints(tableParam.getHeight());

            for (int j=0;j<tableParam.getCols().size();j++) {
                // 创建行的单元格
                Cell cell = _row.createCell(j);

                //判断列是否为设置公式的列,设置公式不根据key生成数据,直接写入公式即可
                if(Utils.notEmpty(tableParam.getCols().get(j).getFormula())){
                    cell.setCellFormula(tableParam.getCols().get(j).getFormula());
                    continue;
                }

                //根据Map的key获取value值
                Object result = data.get(k).get(tableParam.getCols().get(j).getKey());
                setCell(cell,result,tableParam.getCols().get(j));
            }
            currentRow++;
        }
        return currentRow;
    }

    /**
     * 行配置 （标题、下拉）
     * @param workbook workbook
     * @param sheet sheet
     * @param tableParam tableParam
     */
    private static void rowConfig(Workbook workbook, Sheet sheet, TableParam tableParam){
        List<Col> cols = tableParam.getCols();
        //标题配置
        if(tableParam.getCreateHeadRow()){
            /*创建标题行*/
            Row row = sheet.createRow(tableParam.getWriteRow());
            //标题行样式
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            //加粗
            font.setBold(tableParam.getHeadRowStyle().getHeadBold());
            style.setFont(font);
            //居中
            style.setAlignment(tableParam.getHeadRowStyle().getHorizontalAlignment());
            for(int i=0;i<cols.size();i++){
                // 创建行的单元格,也是从0开始
                Cell cell = row.createCell(i);
                // 设置单元格内容
                cell.setCellValue(cols.get(i).getTitle());
                cell.setCellStyle(style);
            }
        }
        //读取下拉菜单配置
        List<DropdownParam> dropdownParams = new ArrayList<DropdownParam>();
        for (int i = 0; i < cols.size(); i++) {
            if(cols.get(i).getDropdownList() != null){
                DropdownParam temp = new DropdownParam(tableParam.getWriteRow(),200,i,i,cols.get(i).getDropdownList());
                dropdownParams.add(temp);
            }

            //设置列宽
            sheet.setColumnWidth(i,tableParam.getCols().get(i).getWidth()*256);
        }
        if(dropdownParams.size() > 0) {
            if (workbook instanceof XSSFWorkbook) {
                setDropDownList((XSSFSheet) sheet, dropdownParams);
            } else if (workbook instanceof HSSFWorkbook) {
                setDropDownList((HSSFSheet) sheet, dropdownParams);
            }
        }
    }

    //默认日期转换格式
    private static final String FORMAT="yyyy/m/d";
    private static final String FORMAT2="yyyy/m/d h:mm:ss;@";
    private static SimpleDateFormat sdfDate= new SimpleDateFormat(FORMAT);
    private static SimpleDateFormat sdfTime= new SimpleDateFormat(FORMAT2);

    /**
     * 单元格cell值设置
     * @param cell cell 列
     * @param result result 值
     * @param col 列s属性对象
     */
    private static void setCell(Cell cell, Object result, Col col){
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
                if (col.getConvertValue() != null) {
                    String val = col.getConvertValue().convert(result);
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
                if (col.getConvertValue() != null) {
                    String val = col.getConvertValue().convert(result);
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
                //String strDate;
                //if (Utils.notEmpty(col.getFormat())) {
                //    SimpleDateFormat s = new SimpleDateFormat(col.getFormat());
                //    strDate = s.format(date);
                //}else{
                //    strDate = sdfTime.format(date);
                //}
                cell.setCellValue(date);
            }else{
                Date empty = null;
                cell.setCellValue(empty);
            }
            if(Utils.notEmpty(col.getFormat())) {
                if(col.getCellStyle() != null && col.getDataFormat() != null) {
                    col.getCellStyle().setDataFormat(col.getDataFormat().getFormat(col.getFormat()));
                }
            }else{
                if(col.getCellStyle() != null && col.getDataFormat() != null) {
                    col.getCellStyle().setDataFormat(col.getDataFormat().getFormat(FORMAT2));
                }
            }
        }else if(result instanceof LocalDate){
            if(result != null) {
                LocalDate localDate = (LocalDate) result;
                //LocalDate转换为Date
                ZoneId zoneId = ZoneId.systemDefault();
                ZonedDateTime zdt = localDate.atStartOfDay(zoneId);
                Date date = Date.from(zdt.toInstant());
                //String strDate = sdfDate.format(date);
                cell.setCellValue(date);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
            if(Utils.notEmpty(col.getFormat())) {
                if(col.getCellStyle() != null && col.getDataFormat() != null) {
                    col.getCellStyle().setDataFormat(col.getDataFormat().getFormat(col.getFormat()));
                }
            }else{
                if(col.getCellStyle() != null && col.getDataFormat() != null) {
                    col.getCellStyle().setDataFormat(col.getDataFormat().getFormat(FORMAT));
                }
            }
        }else if(result instanceof LocalDateTime){
            if(result != null) {
                LocalDateTime localDateTime = (LocalDateTime) result;
                //LocalDateTime转换为Date
                ZoneId zoneId = ZoneId.systemDefault();
                ZonedDateTime zdt =localDateTime.atZone(zoneId);
                Date date = Date.from(zdt.toInstant());
                //String strDate = sdfTime.format(date);
                cell.setCellValue(date);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
            if(Utils.notEmpty(col.getFormat())) {
                if(col.getCellStyle() != null && col.getDataFormat() != null) {
                    col.getCellStyle().setDataFormat(col.getDataFormat().getFormat(col.getFormat()));
                }
            }else{
                if(col.getCellStyle() != null && col.getDataFormat() != null) {
                    col.getCellStyle().setDataFormat(col.getDataFormat().getFormat(FORMAT2));
                }
            }
        }else{
            if(result != null) {
                cell.setCellValue(resultStr);
            }else{
                String empty = null;
                cell.setCellValue(empty);
            }
        }
        //样式
        cell.setCellStyle(col.getCellStyle());
    }

    /**
     * 设置下拉列表
     * @param sheet sheet
     * @param params 下拉列表配置
     */
    private static void setDropDownList(XSSFSheet sheet,List<DropdownParam> params) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        for (DropdownParam param : params) {
            CellRangeAddressList addressList = new CellRangeAddressList(param.getFirstRow(), param.getLastRow(), param.getFirstCol(),param.getLastCol());
            DataValidationConstraint constraint = helper.createExplicitListConstraint(param.getDropdownList());
            DataValidation dataValidation = helper.createValidation(constraint, addressList);
            dataValidation.createErrorBox("数据非法提醒", "数据不规范,请选择下拉列表中的数据");
            //  处理Excel兼容性问题
            if (dataValidation instanceof XSSFDataValidation) {
                dataValidation.setSuppressDropDownArrow(true);
                dataValidation.setShowErrorBox(true);
            } else {
                dataValidation.setSuppressDropDownArrow(false);
            }
            sheet.addValidationData(dataValidation);
        }
    }

    /**
     * 设置下拉列表
     * @param sheet sheet
     * @param params 下拉列表配置
     */
    private static void setDropDownList(HSSFSheet sheet,List<DropdownParam> params) {
        for (DropdownParam param : params) {
            CellRangeAddressList regions = new CellRangeAddressList(param.getFirstRow(), param.getLastRow(), param.getFirstCol(),param.getLastCol());
            //创建下拉列表数据
            DVConstraint constraint = DVConstraint.createExplicitListConstraint(param.getDropdownList());
            //绑定
            HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
            dataValidation.createErrorBox("数据非法提醒", "数据不规范,请选择下拉列表中的数据");
            sheet.addValidationData(dataValidation);
        }

    }

}
