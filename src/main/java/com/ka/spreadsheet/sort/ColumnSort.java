package com.ka.spreadsheet.sort;

public class ColumnSort {
    public enum SortDirection{
        Ascending,Descending
    }
    private int columnIndex;
    private SortDirection sortDirection;

    public ColumnSort(int columnIndex,SortDirection sortDirection){
        this.columnIndex=columnIndex;
        this.sortDirection=sortDirection;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public SortDirection getSortDirection() {
        return sortDirection;
    }
}
