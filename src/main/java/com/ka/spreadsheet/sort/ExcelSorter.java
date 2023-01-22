package com.ka.spreadsheet.sort;

import com.ka.spreadsheet.util.LogUtil;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.commons.io.FilenameUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.List;


public class ExcelSorter {
    private static Logger log = Logger.getLogger(ExcelSorter.class);
    private static String WORKBOOK_FILENAME;
    private static String LOG_FILENAME;
    private static final String instructions = "FilePath: The fully qualified path to the Excel workbook file you want to sort.\n" +
            "LogFile: The name of the log file to use for the application.\n"+
            "SheetSort: A formatted value that indicates the sheet to be sorted, the columns to sort by and the starting row.\n" +
            "    Each sheet sort argument has the following format:\n"+
            "    [Sheet Index:{Column Index!A (Ascending) or D (Descending),...},Start Row Index]\n" +
            "        Here are a few examples of the above format:\n" +
            "        Sort the first sheet by the first column in ascending order and the second column in descending order. Start the sort at row 1.\n" +
            "        [0:{0!A,1!D}:1]\n" +
            "        Sort the second sheet by the first column in ascending order. Start the sort at row 3.\n" +
            "        [1:{0!A}:2]\n" +
            "        Sort the fourth sheet by the first column in ascending order, the second column in descending order and the fifth column in ascending order. Start the sort at row 1.\n" +
            "        [3:{0!A,1!D,4!A}:0]";

    public static void main(String[] args) throws Exception {
        FileInputStream inputStream = null;
        List<SortSheetParams> sortSheetParamsList = new ArrayList<>();
        try {
            //Parse the arguments
            for(int i=0;i< args.length;i++){
                switch(i){
                    case 0:
                        LOG_FILENAME = args[i];
                        LogUtil.initLogging(LOG_FILENAME);
                        LogUtil.initCleanLogging(LOG_FILENAME);
                        break;
                    case 1:
                        WORKBOOK_FILENAME = args[i];
                        inputStream = new FileInputStream(WORKBOOK_FILENAME);
                        break;
                    default:
                        sortSheetParamsList.add(parseArgument(args[i]));
                }
            }
            XSSFWorkbook workbook  = new XSSFWorkbook(inputStream);
            //Sort the workbook sheets
            sortSheetParamsList.forEach(sortSheetParams -> {sortSheet(workbook, sortSheetParams);});
            inputStream.close();
            // Write the output to a new file
            FileOutputStream outputStream = new FileOutputStream(FilenameUtils.getFullPath(WORKBOOK_FILENAME) + "sorted_" +
                    FilenameUtils.getBaseName(WORKBOOK_FILENAME) +
                    "." + FilenameUtils.getExtension(WORKBOOK_FILENAME));
            workbook.write(outputStream);
            outputStream.close();
        }catch(IllegalArgumentException iex){
            log.error("Unable to process the arguments.",iex);
            System.out.println("Unable to process the arguments.");
            System.out.println(instructions);
            System.exit(100);
        }
        catch (RuntimeException rex){
            log.error("An unexpected runtime exception occurred.",rex);
            System.out.println("An unexpected runtime exception occurred.");
            rex.printStackTrace();
            System.exit(200);
        } catch (Exception ex){
            log.error("An unexpected exception occurred.",ex);
            System.exit(300);
        }
        System.exit(0);
    }

    private static void sortSheet(XSSFWorkbook workbook,SortSheetParams params){
        XSSFSheet sheet = workbook.getSheetAt(params.sheetIndex);
        XSSFSheet newSheet = workbook.createSheet("sorted" + '_' + sheet.getSheetName());
        //Collect the rows in an array list
        ArrayList<Row> rows = new ArrayList<>();
        for(Row row:sheet){
            if(row.getRowNum() < params.rowStartIndex)
                continue;
            rows.add(row);
        }
        //Sort the array list by the criteria
        Collections.sort(rows,
                (row1, row2) -> compareRows(params.columnSortList,row2,row1)
        );
        //Write the sorted rows to the new sheet
        int i = 0;
        for(Row row:rows){
            copyRow(workbook,sheet,newSheet,row.getRowNum(),i);
            i++;
        }
        //Remove rows from the original sheet
        for(Row row:rows){
            sheet.removeRow(row);
        }
        //Copy sorted rows back to the original sheet.
        i=params.rowStartIndex;
        for(Row row:newSheet){
            copyRow(workbook,newSheet,sheet,row.getRowNum(),i);
            i++;
        }
        //Delete the new sheet
        workbook.removeSheetAt(workbook.getNumberOfSheets()-1);
    }

    private static SortSheetParams parseArgument(String sheetSortArg){
        SortSheetParams sortSheetParams = new SortSheetParams();
        // Regular expression to match the format
        String pattern = "\\[(\\d+)\\:\\{(.*?)\\}\\:(\\d+)\\]";
        // Compile the regular expression
        Pattern r = Pattern.compile(pattern);
        // Match the input against the regular expression
        Matcher m = r.matcher(sheetSortArg);
        log.info("Parse Argument:");
        if (m.find()) {
            sortSheetParams.sheetIndex = Integer.parseInt(m.group(1));
            sortSheetParams.columnSortList = parseColumnSortArg(m.group(2));
            sortSheetParams.rowStartIndex = Integer.parseInt(m.group(3));
            log.info("....sheetIndex:" + sortSheetParams.sheetIndex);
            log.info("....columnSortList:" + m.group(2));
            log.info("....rowStartIndex:" + sortSheetParams.rowStartIndex);
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


    private static void swapRows(XSSFWorkbook workbook, XSSFSheet sheet,Row row1,Row row2){
        int row2Index = row2.getRowNum();
        int lastRow = sheet.getLastRowNum();
        //Create a row after row2 as placeholder
        sheet.shiftRows(row2.getRowNum() + 1, row2Index == lastRow ? lastRow+1:lastRow,
                1, true, true);
        //Copy row 1 into the new row
        copyRow(workbook,sheet,sheet,row1.getRowNum(), row2.getRowNum() + 1);
        //Remove row 1
        removeRow(sheet, row1.getRowNum());
    }

    private static int compareRows(List<ColumnSort> columnSortList,Row row,Row row2) {
        int ret = 0;
        for (ColumnSort columnSort : columnSortList) {
            Cell cell1 = row.getCell(columnSort.getColumnIndex());
            Cell cell2 = row2.getCell(columnSort.getColumnIndex());
            ret = compareCells(cell1,cell2,columnSort.getSortDirection());
            if(ret != 0) break;
        }
        return ret;
    }

    private static int compareCells(Cell row1Cell, Cell row2Cell, ColumnSort.SortDirection sortDirection) {
        if (row1Cell == null || row2Cell == null) return 0;
        int ret;
        switch(row1Cell.getCellType()){
            case Cell.CELL_TYPE_BOOLEAN:
                boolean bvalue1 = row1Cell.getBooleanCellValue();
                boolean bvalue2 = row2Cell.getBooleanCellValue();
                ret = Boolean.compare(bvalue2,bvalue1);
                break;
            case Cell.CELL_TYPE_STRING:
                String svalue1 = row1Cell.getStringCellValue();
                String svalue2 = row2Cell.getStringCellValue();
                ret = svalue2.compareToIgnoreCase(svalue1);
                break;
            case Cell.CELL_TYPE_NUMERIC:
                Double dvalue1 = row1Cell.getNumericCellValue();
                Double dvalue2 = row2Cell.getNumericCellValue();
                ret = dvalue2.compareTo(dvalue1);
                break;
            default:
                ret = 0;
        }
        return sortDirection == ColumnSort.SortDirection.Ascending ? ret : -ret;
    }

    private static void copyRow(XSSFWorkbook workbook, XSSFSheet srcWorksheet, XSSFSheet destWorksheet, int sourceRowNum, int destinationRowNum) {
        // Get the source / new row
        XSSFRow newRow = destWorksheet.getRow(destinationRowNum);
        XSSFRow sourceRow = srcWorksheet.getRow(sourceRowNum);

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            destWorksheet.shiftRows(destinationRowNum, destWorksheet.getLastRowNum(), 1);
        } else {
            newRow = destWorksheet.createRow(destinationRowNum);
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
        for (int i = 0; i < srcWorksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = srcWorksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                destWorksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }

    /**
     * Remove a row by its index
     * @param sheet an Excel sheet
     * @param rowIndex a 0 based index of the row being removed
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
