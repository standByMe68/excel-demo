package com.standbyme.exceldemo.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtils {

    private static Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    public static Workbook getWorkbook(MultipartFile file) {
        Workbook workbook = null;
        String name = file.getOriginalFilename();
        String suffix = name.substring(name.indexOf("."));

        InputStream inputStream = null;
        logger.info("file suffix name:{}", suffix);
        try {
            inputStream = file.getInputStream();
            if (ExcelContant.XLSX.equals(suffix)) {
                workbook = new XSSFWorkbook(inputStream);
            } else if (ExcelContant.XLS.equals(suffix)) {
                workbook = new HSSFWorkbook(inputStream);
            } else {
                throw new FileSuffixException();
            }
        } catch (IOException e) {
            logger.error("获取文件流失败");
            e.printStackTrace();
        } catch (FileSuffixException e) {
            logger.error("文件后缀名不符合要求");
            e.printStackTrace();
        }
        try {
            inputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    private static Sheet getSheet(Workbook workbook, int pageNum) {
        return workbook.getSheetAt(pageNum);
    }

    /**
     * 获取Excel表格中的数据
     * @param workbook
     * @param columns 列名数组
     * @param headNum 表头行数
     * @return
     */
    public static List<Map<String, String>> getSheetValue(Workbook workbook,String[] columns,int headNum,int pageNum) {
        List<Map<String, String>> sheetValues = new ArrayList<>();
        Sheet sheet = getSheet(workbook, pageNum);
        if (sheet != null) {
            int lastRowNum = sheet.getPhysicalNumberOfRows()-headNum;
            for (int i = headNum; i < lastRowNum; i++) {
                Map<String, String> rowValues = new HashMap<>();
                Row row = sheet.getRow(i);
                if (row != null) {
                    rowValues.put("行数", i + "");
                    for (int j = 0; j < columns.length; j++) {
                        Cell cell = row.getCell(j);
                        if (cell != null) {
                            String valueByCell = getValueByCell(cell);
                            if (!StringUtils.isEmpty(valueByCell)) {
                                rowValues.put(columns[j], valueByCell);
                            }
                        } else {
                            logger.info("第{}行第{}个格子为空",i,j);
                        }
                    }
                    if (rowValues.size() >= 2) {
                        sheetValues.add(rowValues);
                    }
                } else {
                    logger.info("第{}行数据为空",i);
                }
            }
        } else {
            logger.error("sheet is null");
        }
        return sheetValues;
    }

    private static String getValueByCell(Cell cell) {
        String value = "";
        if (cell != null) {
            int cellType = cell.getCellType();
            switch (cellType) {
                case Cell.CELL_TYPE_NUMERIC:
                    value = String.valueOf(cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        Date dateCellValue = cell.getDateCellValue();
                        value = simpleDateFormat.format(dateCellValue);
                    } else {
                        value = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                case Cell.CELL_TYPE_BLANK:
                    value = "";
                case Cell.CELL_TYPE_BOOLEAN:
                    value = String.valueOf(cell.getBooleanCellValue());
                default:
                    value = "";
            }
        }else {
            logger.info("cell is null");
        }
        return value;
    }



}
