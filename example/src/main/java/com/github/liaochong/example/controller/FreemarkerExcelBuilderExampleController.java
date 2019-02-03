package com.github.liaochong.example.controller;

import com.github.liaochong.html2excel.core.ExcelBuilder;
import com.github.liaochong.html2excel.core.FreemarkerExcelBuilder;
import org.apache.commons.codec.CharEncoding;
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
    public void build(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("data", this.getData());

        Workbook workbook = excelBuilder.template("/templates/freemarkerToExcelExample.ftl").build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("freemarker_excel.xlsx".getBytes()));
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
    @GetMapping("/freemarker/defaultStyle/example")
    public void buildWithDefaultStyle(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("data", this.getData());

        Workbook workbook = excelBuilder.template("/templates/freemarkerToExcelExample.ftl").useDefaultStyle().build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("freemarker_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<Integer> getData() {
        List<Integer> data = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            data.add(i);
        }
        return data;
    }
}
