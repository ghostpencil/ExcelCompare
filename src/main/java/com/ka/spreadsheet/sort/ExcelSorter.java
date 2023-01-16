package com.ka.spreadsheet.sort;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.io.FilenameUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.List;


public class ExcelSorter {

    private static String WORKBOOK_FILENAME;
/*
    FilePath:
    SheetSort: A formatted value that indicates the sheet to be sorted, the columns to sort by and the starting row.
            [Sheet Index:{Column Index!A (Ascending) or D (Descending),...},Start Row Index]
            Here are a few examples of the above format:
            Sort the first sheet by the first column in ascending order and the second column in descending order. Start the sort at row 1.
            [0:{0!A,1!D}:1]
            Sort the second sheet by the first column in ascending order. Start the sort at row 3.
            [1:{0!A}:2]
            Sort the fourth sheet by the first column in ascending order, the second column in descending order and the fifth column in ascending order. Start the sort at row 1.
            [3:{0!A,1!D,4!A}:0]
 */

    public static void main(String[] args) throws Exception {
        WORKBOOK_FILENAME = args[0];
        SortSheetParams sortSheetParams1 = parseArgument(args[1]);
        FileInputStream inputStream = new FileInputStream(WORKBOOK_FILENAME);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        try {
            sortSheet(workbook, sortSheetParams1);
            inputStream.close();
            // Write the output to a new file
            FileOutputStream outputStream = new FileOutputStream(FilenameUtils.getFullPath(WORKBOOK_FILENAME) + "sorted_" +
                    FilenameUtils.getBaseName(WORKBOOK_FILENAME) +
                    "." + FilenameUtils.getExtension(WORKBOOK_FILENAME));
            workbook.write(outputStream);
            outputStream.close();
        } catch (RuntimeException rex){
           rex.printStackTrace();
        } catch (Exception ex){
            ex.printStackTrace();
        }
    }

    private static SortSheetParams parseArgument(String sheetSortArg){
        SortSheetParams sortSheetParams = new SortSheetParams();
        // Regular expression to match the format
        String pattern = "\\[(\\d+)\\:\\{(.*?)\\}\\:(\\d+)\\]";
        // Compile the regular expression
        Pattern r = Pattern.compile(pattern);
        // Match the input against the regular expression
        Matcher m = r.matcher(sheetSortArg);
        System.out.println("Parse Argument:");
        if (m.find()) {
            sortSheetParams.sheetIndex = Integer.parseInt(m.group(1));
            sortSheetParams.columnSortList = parseColumnSortArg(m.group(2));
            sortSheetParams.rowStartIndex = Integer.parseInt(m.group(3));
            System.out.println("....sheetIndex:" + sortSheetParams.sheetIndex);
            System.out.println("....columnSortList:" + m.group(2));
            System.out.println("....rowStartIndex:" + sortSheetParams.rowStartIndex);
        }
        return sortSheetParams;
    }

    private static List<ColumnSort> parseColumnSortArg(String columnSortArg){
        List<ColumnSort> columnSortList = new ArrayList<>();
        // Split the second value string by ","
        String[] secondValuePairs = columnSortArg.split(",");
        ColumnSort colSort;
        for (String pair : secondValuePairs) {
            // Split the pair by "!"
            String[] keyValue = pair.split("!");
            int key = Integer.parseInt(keyValue[0]);
            String value = keyValue[1];
            colSort = new ColumnSort(key,value.compareToIgnoreCase("A") == 0 ? ColumnSort.SortDirection.Ascending: ColumnSort.SortDirection.Descending);
            columnSortList.add(colSort);
        }
        return columnSortList;
    }

    /**
     * Sorts (A-Z) rows by String column
     * @param workbook - Workbook to sort
     * @param params - Sheet sort parameters that include the row start index and the list of columns to sort by
     */
    private static void sortSheet(XSSFWorkbook workbook, SortSheetParams params) {
        boolean sorting = true;
        XSSFSheet sheet = workbook.getSheetAt(params.sheetIndex);
        int lastRow = sheet.getLastRowNum();
        while (sorting) {
            sorting = false;
            for(int i = params.rowStartIndex;i < lastRow;i++) {
                System.out.println(i);
                Row row = sheet.getRow(i);
                if (row == null) continue;
                // end if this is last row
                if (lastRow==row.getRowNum()) break;
                Row row2 = sheet.getRow(row.getRowNum()+1);
                if (row2 == null) continue;
                //compare cell from current row and next row - and switch if secondValue should be before first
                if (compareRows(params.columnSortList,row,row2)) {
                    System.out.println("ROW:" + row.getRowNum());
                    swapRows(workbook,sheet,row,row2);
                    sorting = true;
                }
            }
        }
    }

    private static void swapRows(XSSFWorkbook workbook, XSSFSheet sheet,Row row1,Row row2){
        int lastRow = sheet.getLastRowNum();
        //Create a row after row2 as placeholder
        sheet.shiftRows(row2.getRowNum()+1,lastRow,1,true,true);
        //Copy row 1 into the new row
        copyRow(workbook,sheet,row1.getRowNum(),row2.getRowNum()+1);
        //Remove row 1
        removeRow(sheet,row1.getRowNum());
    }

    private static boolean compareRows(List<ColumnSort> columnSortList,Row row,Row row2) {
        boolean ret = false;
        for (ColumnSort columnSort : columnSortList) {
            Cell cell1 = row.getCell(columnSort.getColumnIndex());
            Cell cell2 = row2.getCell(columnSort.getColumnIndex());
            ret = compareCells(cell1,cell2,columnSort.getSortDirection());
            if(ret) break;
        }
        return ret;
    }

    private static boolean compareCells(Cell row1Cell, Cell row2Cell, ColumnSort.SortDirection sortDirection) {
        if (row1Cell == null || row2Cell == null) return false;
        boolean ret;
        switch(row1Cell.getCellType()){
            case Cell.CELL_TYPE_BOOLEAN:
                boolean bvalue1 = row1Cell.getBooleanCellValue();
                boolean bvalue2 = row2Cell.getBooleanCellValue();
                ret = sortDirection == ColumnSort.SortDirection.Ascending ?
                        Boolean.compare(bvalue2,bvalue1) <0:
                        Boolean.compare(bvalue2,bvalue1) >0;
                break;
            case Cell.CELL_TYPE_STRING:
                String svalue1 = row1Cell.getStringCellValue();
                String svalue2 = row2Cell.getStringCellValue();
                ret = sortDirection == ColumnSort.SortDirection.Ascending ?
                        svalue2.compareToIgnoreCase(svalue1) < 0:
                        svalue2.compareToIgnoreCase(svalue1) > 0;
                break;
            case Cell.CELL_TYPE_NUMERIC:
                Double dvalue1 = row1Cell.getNumericCellValue();
                Double dvalue2 = row2Cell.getNumericCellValue();
                ret = sortDirection == ColumnSort.SortDirection.Ascending ?
                        dvalue2.compareTo(dvalue1) < 0:
                        dvalue2.compareTo(dvalue1) > 0;
                break;
            default:
                ret = false;
        }
        return ret;
    }

    private static void copyRow(XSSFWorkbook workbook, XSSFSheet worksheet, int sourceRowNum, int destinationRowNum) {
        // Get the source / new row
        XSSFRow newRow = worksheet.getRow(destinationRowNum);
        XSSFRow sourceRow = worksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(destinationRowNum);
        }

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            XSSFCell oldCell = sourceRow.getCell(i);
            XSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            XSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }

        // If there are are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }

    /**
     * Remove a row by its index
     * @param sheet a Excel sheet
     * @param rowIndex a 0 based index of removing row
     */
    public static void removeRow(XSSFSheet sheet, int rowIndex) {
        int lastRowNum=sheet.getLastRowNum();
        if(rowIndex>=0&&rowIndex<lastRowNum){
            sheet.shiftRows(rowIndex+1,lastRowNum, -1);
        }
        if(rowIndex==lastRowNum){
            XSSFRow removingRow=sheet.getRow(rowIndex);
            if(removingRow!=null){
                sheet.removeRow(removingRow);
            }
        }
    }

    private static class SortSheetParams{
        int sheetIndex;
        List<ColumnSort> columnSortList;
        int rowStartIndex;
    }
}
