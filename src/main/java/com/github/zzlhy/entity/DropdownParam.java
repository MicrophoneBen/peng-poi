package com.github.zzlhy.entity;

/**
 * 下拉列表配置
 * Created by Administrator on 2018-11-6.
 */
public class DropdownParam extends Position{

    //下拉列表数组
    private String[] dropdownList;

    public DropdownParam() {
    }

    public DropdownParam(int firstRow, int lastRow, int firstCol, int lastCol, String[] dropdownList) {
        super(firstRow, lastRow, firstCol, lastCol);
        this.dropdownList = dropdownList;
    }

    public DropdownParam of(int firstRow, int lastRow, int firstCol, int lastCol, String[] dropdownList) {
        return new DropdownParam(firstRow,lastRow,firstCol,lastCol,dropdownList);
    }

    public String[] getDropdownList() {
        return dropdownList;
    }

    public void setDropdownList(String[] dropdownList) {
        this.dropdownList = dropdownList;
    }
}
