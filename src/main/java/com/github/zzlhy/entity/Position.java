package com.github.zzlhy.entity;

/**
 * 表示Excel的某个区域位置
 * Created by Administrator on 2018-11-6.
 */
public class Position {

    //开始行
    private int firstRow;

    //结束行
    private int lastRow;

    //开始列
    private int firstCol;

    //结束列
    private int lastCol;

    public Position() {
    }

    public Position(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public void setFirstCol(int firstCol) {
        this.firstCol = firstCol;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public int getLastCol() {
        return lastCol;
    }

    public void setLastCol(int lastCol) {
        this.lastCol = lastCol;
    }
}
