package com.standbyme.exceldemo.controller;

import com.standbyme.exceldemo.util.ExcelUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.util.List;

public class TestAPiMain {

    public static void main(String[] args) {
        File file = new File("D:\\临时文件夹\\报警定义导入模板20210220(1).xlsx");
        Workbook workbook = ExcelUtils.getWorkbook(file);
        int activeSheetIndex = workbook.getActiveSheetIndex();
        System.out.println("activeSheetIndex = " + activeSheetIndex);

        List<? extends Name> allNames = workbook.getAllNames();
        System.out.println("allNames = " + allNames);

        SpreadsheetVersion spreadsheetVersion = workbook.getSpreadsheetVersion();
        System.out.println("spreadsheetVersion = " + spreadsheetVersion);

        System.out.println("========================sheet API测试============================");
        Sheet sheetAt = workbook.getSheetAt(0);
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        System.out.println("physicalNumberOfRows = " + physicalNumberOfRows);

        int lastRowNum = sheetAt.getLastRowNum();
        System.out.println("lastRowNum = " + lastRowNum);

        //和创建sheet同一个对象
        Workbook workbook1 = sheetAt.getWorkbook();
        System.out.println(workbook1.equals(workbook));
        System.out.println(workbook1 == workbook);

        boolean autobreaks = sheetAt.getAutobreaks();
        System.out.println("autobreaks = " + autobreaks);

        CellAddress activeCell = sheetAt.getActiveCell();
        System.out.println("activeCell = " + activeCell);
        int row = activeCell.getRow();
        System.out.println("row = " + row);
        int column = activeCell.getColumn();
        System.out.println("column = " + column);

        List<CellRangeAddress> mergedRegions = sheetAt.getMergedRegions();
        System.out.println("mergedRegions = " + mergedRegions);

        System.out.println("========================row cell API测试============================");
        Row row1 = sheetAt.getRow(0);
        Cell cell = row1.getCell(0);
//        System.out.println(cell.getStringCellValue());

        short firstCellNum = row1.getFirstCellNum();
        System.out.println("firstCellNum = " + firstCellNum);

        short height = row1.getHeight();
        System.out.println("height = " + height);

        int physicalNumberOfCells = row1.getPhysicalNumberOfCells();
        System.out.println("physicalNumberOfCells = " + physicalNumberOfCells);

        //数组公式
        //CellRangeAddress arrayFormulaRange = cell.getArrayFormulaRange();
    }

}
