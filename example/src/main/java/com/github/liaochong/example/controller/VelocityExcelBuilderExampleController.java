package com.github.liaochong.example.controller;

import com.github.liaochong.example.pojo.Product;
import com.github.liaochong.myexcel.core.ExcelBuilder;
import com.github.liaochong.myexcel.core.VelocityExcelBuilder;
import com.github.liaochong.myexcel.core.WorkbookType;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.utils.AttachmentExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author gaokai
 * @version 1.0
 */
@RestController
public class VelocityExcelBuilderExampleController {

    /**
     * use non-default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/velocity/example")
    public void build(HttpServletResponse response) throws IOException {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();
            Workbook workbook = excelBuilder.template("/templates/velocityToExcelExample.vm").build(dataMap);
            AttachmentExportUtil.export(workbook, "velocity_excel", response);
        }
    }

    /**
     * use default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/velocity/defaultStyle/example")
    public void buildWithDefaultStyle(HttpServletResponse response) throws IOException {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();

            Workbook workbook = excelBuilder
                    .template("/templates/velocityToExcelExample.vm")
                    .useDefaultStyle()
                    .build(dataMap);
            AttachmentExportUtil.export(workbook, "velocity_excel", response);
        }
    }

    /**
     * build .xls excel
     *
     * @param response response
     */
    @GetMapping("/velocity/xls/example")
    public void buildWithXLS(HttpServletResponse response) throws IOException {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();

            Workbook workbook = excelBuilder
                    .template("/templates/velocityToExcelExample.vm")
                    .workbookType(WorkbookType.XLS)
                    .useDefaultStyle()
                    .build(dataMap);
            AttachmentExportUtil.export(workbook, "velocity_excel", response);
        }
    }

    /**
     * build .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/velocity/xlsx/example")
    public void buildWithXLSX(HttpServletResponse response) throws IOException {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();

            Workbook workbook = excelBuilder
                    .template("/templates/velocityToExcelExample.vm")
                    .workbookType(WorkbookType.XLSX)
                    .useDefaultStyle()
                    .build(dataMap);
            AttachmentExportUtil.export(workbook, "velocity_excel", response);
        }
    }

    /**
     * build .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/velocity/sxlsx/example")
    public void buildWithSXLSX(HttpServletResponse response) throws IOException {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();

            Workbook workbook = excelBuilder
                    .template("/templates/velocityToExcelExample.vm")
                    .workbookType(WorkbookType.SXLSX)
                    .useDefaultStyle()
                    .build(dataMap);
            AttachmentExportUtil.export(workbook, "velocity_excel", response);
        }
    }

    /**
     * encrypt .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/velocity/encrypt/example")
    public void buildWithEncrypt(HttpServletResponse response) throws Exception {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();

            Workbook workbook = excelBuilder
                    .template("/templates/velocityToExcelExample.vm")
                    .workbookType(WorkbookType.SXLSX)
                    .useDefaultStyle()
                    .build(dataMap);
            AttachmentExportUtil.encryptExport(workbook, "velocity_excel", response, "123456");
        }
    }

    /**
     * encrypt .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/velocity/autoWidth/example")
    public void buildWithAutoWidth(HttpServletResponse response) throws Exception {
        try (ExcelBuilder excelBuilder = new VelocityExcelBuilder()) {
            Map<String, Object> dataMap = this.getDataMap();

            Workbook workbook = excelBuilder
                    .template("/templates/velocityToExcelExample.vm")
                    .useDefaultStyle()
                    .autoWidthStrategy(AutoWidthStrategy.AUTO_WIDTH)
                    .build(dataMap);
            AttachmentExportUtil.export(workbook, "velocity_excel", response);
        }
    }

    private Map<String, Object> getDataMap() {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "velocity_excel_example");

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
