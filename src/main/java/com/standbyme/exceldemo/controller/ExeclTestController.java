package com.standbyme.exceldemo.controller;

import com.standbyme.exceldemo.util.ExcelUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/excel")
public class ExeclTestController {

    @ResponseBody
    @PostMapping("/sheetValue")
    public String upload(@RequestParam("file") MultipartFile file) {
        String[] columns = new String[]{"123", "234", "456"};
        Workbook workbook = ExcelUtils.getWorkbook(file);
        List<Map<String, String>> sheetValue = ExcelUtils.getSheetValue(workbook, columns, 3, 0);
        ExcelUtils.setDateInExcel("D:\\临时文件夹\\报警定义导入模板20210220(1).xlsx",0,sheetValue,4,columns);
        return "hello word";
    }

    @RequestMapping("/test")
    public String test() {
        return "index";
    }


}
