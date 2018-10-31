package com.github.zzlhy.entity;

import java.util.List;

/**
 * Excel表格参数配置
 * Created by Administrator on 2017-11-29.
 */
public class TableParam {

    //Sheet名称
    private String sheetName="Sheet";

    //每个sheet允许的数据行数,超过此行数会创建新的sheet (xlsx最大值为1048576)
    private int sheetDataTotal = 500000;

    //导出时开始写入的起始行,默认从0开始
    private int startRow=0;

    //导入时,起始读取行
    private int readRow=1;

    //行高度
    private float height = 15;

    //是否创建标题行
    private boolean createHeadRow=true;

    //标题行设置
    private HeadRowStyle headRowStyle=new HeadRowStyle();

    //导出Excel类型xls/xlsx   默认为xlsx
    private ExcelType excelType=ExcelType.XLSX;

    //冻结列设置
    //固定行的列号 是从1开始
    private int freezeColSplit = 0;
    //固定行的行号
    private int freezeRowSplit = 0;

    //列属性数组 是从1开始
    private List<Col> cols;

    public static TableParam of(){
        return new TableParam();
    }

    public static TableParam of(List<Col> cols){
        return of().setCols(cols);
    }

    public static TableParam of(List<Col> cols,ExcelType excelType){
        return of().setCols(cols).setExcelType(excelType);
    }

    public static TableParam of(List<Col> cols,int freezeRowSplit){
        return of().setCols(cols).setFreezeRowSplit(freezeRowSplit);
    }

    public static TableParam of(List<Col> cols,int freezeColSplit, int freezeRowSplit){
        return of(cols,freezeRowSplit).setFreezeColSplit(freezeColSplit);
    }

    public static TableParam of(List<Col> cols, int freezeRowSplit,ExcelType excelType){
        return of(cols,freezeRowSplit).setExcelType(excelType);
    }

    public static TableParam of(List<Col> cols,int freezeColSplit, int freezeRowSplit,ExcelType excelType){
        return of(cols,freezeColSplit,freezeRowSplit).setExcelType(excelType);
    }

    public static TableParam of(int sheetDataTotal, List<Col> cols) {
        return of(cols).setSheetDataTotal(sheetDataTotal);
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getSheetDataTotal() {
        return sheetDataTotal;
    }

    public TableParam setSheetDataTotal(int sheetDataTotal) {
        this.sheetDataTotal = sheetDataTotal;
        return this;
    }

    public TableParam setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public int getStartRow() {
        return startRow;
    }

    public TableParam setStartRow(int startRow) {
        this.startRow = startRow;
        return this;
    }

    public List<Col> getCols() {
        return cols;
    }

    public TableParam setCols(List<Col> cols) {
        this.cols = cols;
        return this;
    }

    public float getHeight() {
        return height;
    }

    public TableParam setHeight(float height) {
        this.height = height;
        return this;
    }

    public boolean getCreateHeadRow() {
        return createHeadRow;
    }

    public TableParam setCreateHeadRow(boolean createHeadRow) {
        this.createHeadRow = createHeadRow;
        return this;
    }

    public HeadRowStyle getHeadRowStyle() {
        return headRowStyle;
    }

    public TableParam setHeadRowStyle(HeadRowStyle headRowStyle) {
        this.headRowStyle = headRowStyle;
        return this;
    }

    public int getReadRow() {
        return readRow;
    }

    public TableParam setReadRow(int readRow) {
        this.readRow = readRow;
        return this;
    }

    public ExcelType getExcelType() {
        return excelType;
    }

    public TableParam setExcelType(ExcelType excelType) {
        this.excelType = excelType;
        return this;
    }

    public int getFreezeColSplit() {
        return freezeColSplit;
    }

    public TableParam setFreezeColSplit(int freezeColSplit) {
        this.freezeColSplit = freezeColSplit;
        return this;
    }

    public int getFreezeRowSplit() {
        return freezeRowSplit;
    }

    public TableParam setFreezeRowSplit(int freezeRowSplit) {
        this.freezeRowSplit = freezeRowSplit;
        return this;
    }
}
