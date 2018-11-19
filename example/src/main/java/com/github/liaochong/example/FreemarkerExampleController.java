package com.github.liaochong.example;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletResponse;

import com.github.liaochong.html2excel.core.DefaultExcelBuilder;
import com.github.liaochong.html2excel.core.ExcelBuilder;
import com.github.liaochong.html2excel.core.FreemarkerExcelBuilder;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;


/**
 * @author liaochong
 * @version 1.0
 */
@RestController
public class FreemarkerExampleController {

    @GetMapping("/freemarker/build")
    public void build(HttpServletResponse response) {
        List<Test> dataList = new ArrayList<>();
        for (int i = 0; i < 1000; i++) {
            Test t1 = new Test();
            t1.setName("liaochong");
            t1.setAge(15);
            dataList.add(t1);
        }

        List<String> order = new ArrayList<>();
        order.add("age");
        order.add("name");


//        Workbook workbook = DefaultExcelBuilder.getInstance().sheetName("测试").fieldDisplayOrder(order).build(dataList);


        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Workbook workbook = excelBuilder.template("/templates/freemarker_template.ftl").useDefaultStyle().build(new HashMap<>());

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("freemarker_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Data
    @FieldDefaults(level = AccessLevel.PRIVATE)
    public static class Test {

        String name;

        Integer age;
    }
}
