package com.github.liaochong.example;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.github.liaochong.html2excel.core.ExcelBuilder;
import com.github.liaochong.html2excel.core.FreemarkerExcelBuilder;

/**
 * @author liaochong
 * @version 1.0
 */
@RestController
public class FreemarkerExampleController {

    @GetMapping("/freemarker/build")
    public void build(HttpServletResponse response) {
        Map<String, Object> data = getData();

        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Workbook workbook = excelBuilder.template("/templates/freemarker_template.ftl").build(data);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("freemarker_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private Map<String, Object> getData() {
        Map<String, Object> data = new HashMap<>();
        for (int i = 1; i <= 11; i++) {
            data.put("n_"+String.valueOf(i), i);
        }
        return data;
    }
}
