package com.github.liaochong.example;

import com.github.liaochong.html2excel.core.DefaultExcelBuilder;
import com.github.liaochong.html2excel.core.ExcelBuilder;
import com.github.liaochong.html2excel.core.FreemarkerExcelBuilder;
import com.github.liaochong.html2excel.core.WorkbookType;
import com.github.liaochong.html2excel.core.annotation.ExcelColumn;
import com.github.liaochong.html2excel.core.annotation.ExcelTable;
import com.github.liaochong.html2excel.core.annotation.ExcludeColumn;
import lombok.AccessLevel;
import lombok.experimental.FieldDefaults;
import org.apache.commons.codec.CharEncoding;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;


/**
 * @author liaochong
 * @version 1.0
 */
@RestController
public class FreemarkerExampleController {
    /**
     * use default excel builder
     *
     * @param response response
     */
    @GetMapping("/freemarker/default/build")
    public void defaultBuild(HttpServletResponse response) {
        List<Child> dataList = new ArrayList<>();
        for (int i = 0; i < 500000; i++) {
            Child child = new Child();
            child.setName("liaochong");
            child.setAge(i);
            child.setSex(1);
            child.setIndex(i);
            dataList.add(child);
        }
        Workbook workbook = DefaultExcelBuilder.getInstance().build(dataList);

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("freemarker_excel.xls".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * use non-default-style excel builder
     *
     * @param response response
     */
    @GetMapping("/freemarker/build")
    public void build(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Workbook workbook = excelBuilder.template("/templates/freemarker_template.ftl").build(new HashMap<>());

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
    @GetMapping("/freemarker/default_style/build")
    public void buildWithDefaultStyle(HttpServletResponse response) {
        ExcelBuilder excelBuilder = new FreemarkerExcelBuilder();
        Workbook workbook = excelBuilder.template("/templates/freemarker_template.ftl").workbookType(WorkbookType.SXLSX).useDefaultStyle().build(new HashMap<>());

        response.setCharacterEncoding(CharEncoding.UTF_8);
        response.addHeader("Content-Disposition", "attachment;filename=" + new String("freemarker_excel.xlsx".getBytes()));
        try {
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    @ExcelTable(workbookType = WorkbookType.SXLSX, rowAccessWindowSize = 100, sheetName = "测试")
    public static class Child extends Parent {
        @ExcelColumn(order = 3, title = "姓名")
        private String name;

        @ExcludeColumn
        private Integer age;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }
    }

    @FieldDefaults(level = AccessLevel.PRIVATE)
    public static class Parent {

        @ExcelColumn(title = "性别")
        private Integer sex;

        //        @ExcelColumn(order = -1, title = "index")
        @ExcludeColumn
        private Integer index;

        public Integer getSex() {
            return sex;
        }

        public void setSex(Integer sex) {
            this.sex = sex;
        }

        public Integer getIndex() {
            return index;
        }

        public void setIndex(Integer index) {
            this.index = index;
        }
    }
}
