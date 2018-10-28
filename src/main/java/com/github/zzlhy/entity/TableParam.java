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

    public TableParam() {
    }

    public TableParam(List<Col> cols) {
        this.cols = cols;
    }

    public TableParam(int freezeColSplit, int freezeRowSplit, List<Col> cols) {
        this.freezeColSplit = freezeColSplit;
        this.freezeRowSplit = freezeRowSplit;
        this.cols = cols;
    }

    public TableParam(int sheetDataTotal, List<Col> cols) {
        this.sheetDataTotal = sheetDataTotal;
        this.cols = cols;
    }

    public TableParam(String sheetName, Integer startRow, float height) {
        this.sheetName = sheetName;
        this.startRow = startRow;
        this.height = height;
    }

    public TableParam(ExcelType excelType) {
        this.excelType = excelType;
    }

    public TableParam(ExcelType excelType, List<Col> cols) {
        this.excelType = excelType;
        this.cols = cols;
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getSheetDataTotal() {
        return sheetDataTotal;
    }

    public void setSheetDataTotal(int sheetDataTotal) {
        this.sheetDataTotal = sheetDataTotal;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public List<Col> getCols() {
        return cols;
    }

    public void setCols(List<Col> cols) {
        this.cols = cols;
    }

    public float getHeight() {
        return height;
    }

    public void setHeight(float height) {
        this.height = height;
    }

    public boolean getCreateHeadRow() {
        return createHeadRow;
    }

    public void setCreateHeadRow(boolean createHeadRow) {
        this.createHeadRow = createHeadRow;
    }

    public HeadRowStyle getHeadRowStyle() {
        return headRowStyle;
    }

    public void setHeadRowStyle(HeadRowStyle headRowStyle) {
        this.headRowStyle = headRowStyle;
    }

    public int getReadRow() {
        return readRow;
    }

    public void setReadRow(int readRow) {
        this.readRow = readRow;
    }

    public ExcelType getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelType excelType) {
        this.excelType = excelType;
    }

    public int getFreezeColSplit() {
        return freezeColSplit;
    }

    public void setFreezeColSplit(int freezeColSplit) {
        this.freezeColSplit = freezeColSplit;
    }

    public int getFreezeRowSplit() {
        return freezeRowSplit;
    }

    public void setFreezeRowSplit(int freezeRowSplit) {
        this.freezeRowSplit = freezeRowSplit;
    }
}
