package com.github.liaochong.example.controller;

import com.github.liaochong.html2excel.core.BeetlExcelBuilder;
import com.github.liaochong.html2excel.core.ExcelBuilder;
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
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("data", this.getData());

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
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("data", this.getData());

        Workbook workbook = excelBuilder.template("/templates/beetlToExcelExample.btl").useDefaultStyle().build(dataMap);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("beetl_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<Integer> getData() {
        List<Integer> data = new ArrayList<>();
        for (int i = 0; i < 1000; i++) {
            data.add(i);
        }
        return data;
    }
}
