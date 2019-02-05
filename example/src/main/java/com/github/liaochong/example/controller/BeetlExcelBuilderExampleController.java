package com.github.liaochong.example.controller;

import com.github.liaochong.example.pojo.Product;
import com.github.liaochong.html2excel.core.BeetlExcelBuilder;
import com.github.liaochong.html2excel.core.ExcelBuilder;
import com.github.liaochong.html2excel.core.WorkbookType;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
@Controller
public class BeetlExcelBuilderExampleController {

    /**
     * use non-default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/beetl/example")
    public void build(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder.template("/templates/beetlToExcelExample.btl").build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("beetl_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * use default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/beetl/defaultStyle/example")
    public void buildWithDefaultStyle(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder.template("/templates/beetlToExcelExample.btl").useDefaultStyle().build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("beetl_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * build .xls excel
     *
     * @param response response
     */
    @GetMapping("/beetl/xls/example")
    public void buildWithXLS(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/beetlToExcelExample.btl")
                .workbookType(WorkbookType.XLS)
                .useDefaultStyle()
                .build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("beetl_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * build .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/beetl/xlsx/example")
    public void buildWithXLSX(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/beetlToExcelExample.btl")
                .workbookType(WorkbookType.XLSX)
                .useDefaultStyle()
                .build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("beetl_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * build .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/beetl/sxlsx/example")
    public void buildWithSXLSX(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new BeetlExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/beetlToExcelExample.btl")
                .workbookType(WorkbookType.SXLSX)
                .useDefaultStyle()
                .build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("beetl_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Map<String, Object> getDataMap() {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "beetl_excel_example");

        List<String> titles = new ArrayList<>();
        titles.add("Category");
        titles.add("Product Name");
        titles.add("Count");
        dataMap.put("titles", titles);

        List<Product> data = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Product product = new Product();
            if (i % 2 == 0) {
                product.setCategory("蔬菜");
                product.setName("小白菜");
                product.setCount(100);
            } else {
                product.setCategory("电子产品");
                product.setName("ipad");
                product.setCount(999);
            }
            data.add(product);
        }
        dataMap.put("data", data);
        return dataMap;
    }
}
