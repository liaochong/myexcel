package com.github.liaochong.example.controller;

import com.github.liaochong.example.pojo.Product;
import com.github.liaochong.myexcel.core.ExcelBuilder;
import com.github.liaochong.myexcel.core.FreemarkerExcelBuilder;
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
 * @author liaochong
 * @version 1.0
 */
@RestController
public class FreemarkerExcelBuilderExampleController {

    /**
     * use non-default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/freemarker/example")
    public void build(HttpServletResponse response) throws IOException {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder.template("/templates/freemarkerToExcelExample.ftl").build(dataMap);
        AttachmentExportUtil.export(workbook, "freemarker_excel", response);
    }

    /**
     * use default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/freemarker/defaultStyle/example")
    public void buildWithDefaultStyle(HttpServletResponse response) throws IOException {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/freemarkerToExcelExample.ftl")
                .useDefaultStyle()
                .build(dataMap);
        AttachmentExportUtil.export(workbook, "freemarker_excel", response);
    }

    /**
     * build .xls excel
     *
     * @param response response
     */
    @GetMapping("/freemarker/xls/example")
    public void buildWithXLS(HttpServletResponse response) throws IOException {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/freemarkerToExcelExample.ftl")
                .workbookType(WorkbookType.XLS)
                .useDefaultStyle()
                .build(dataMap);
        AttachmentExportUtil.export(workbook, "freemarker_excel", response);
    }

    /**
     * build .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/freemarker/xlsx/example")
    public void buildWithXLSX(HttpServletResponse response) throws IOException {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/freemarkerToExcelExample.ftl")
                .workbookType(WorkbookType.XLSX)
                .useDefaultStyle()
                .build(dataMap);
        AttachmentExportUtil.export(workbook, "freemarker_excel", response);
    }

    /**
     * build .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/freemarker/sxlsx/example")
    public void buildWithSXLSX(HttpServletResponse response) throws IOException {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/freemarkerToExcelExample.ftl")
                .workbookType(WorkbookType.SXLSX)
                .useDefaultStyle()
                .build(dataMap);
        AttachmentExportUtil.export(workbook, "freemarker_excel", response);
    }

    /**
     * encrypt .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/freemarker/encrypt/example")
    public void buildWithEncrypt(HttpServletResponse response) throws Exception {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/freemarkerToExcelExample.ftl")
                .workbookType(WorkbookType.SXLSX)
                .useDefaultStyle()
                .build(dataMap);
        AttachmentExportUtil.encryptExport(workbook, "freemarker_excel", response, "123456");
    }

    /**
     * encrypt .xlsx excel
     *
     * @param response response
     */
    @GetMapping("/freemarker/autoWidth/example")
    public void buildWithAutoWidth(HttpServletResponse response) throws Exception {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = this.getDataMap();

        Workbook workbook = excelBuilder
                .template("/templates/freemarkerToExcelExample.ftl")
                .useDefaultStyle()
                .autoWidthStrategy(AutoWidthStrategy.AUTO_WIDTH)
                .build(dataMap);
        AttachmentExportUtil.export(workbook, "freemarker_excel", response);
    }

    private Map<String, Object> getDataMap() {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("sheetName", "freemarker_excel_example");

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
