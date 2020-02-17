package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CommonPeople;
import com.github.liaochong.myexcel.core.pojo.CustomStylePeople;
import com.github.liaochong.myexcel.core.pojo.Formula;
import com.github.liaochong.myexcel.core.pojo.OddEvenStylePeople;
import com.github.liaochong.myexcel.core.pojo.WidthPeople;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
class DefaultStreamExcelBuilderTest extends BasicTest {

    @Test
    void mapBuild() throws Exception {
        List<Map> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, String> obj = new HashMap<>();
            obj.put("1", "1");
            obj.put("2", "2");
            obj.put("3", "3");
            obj.put("4", "4");
            list.add(obj);
        }
        List<String> titles = new ArrayList<>();
        titles.add("1->1.1");
        titles.add("2");
        titles.add("3");
        titles.add("4");

        Workbook workbook = DefaultExcelBuilder.of(Map.class).fieldDisplayOrder(titles).titles(titles).build(list);
//        workbook = DefaultExcelBuilder.of(Map.class).fieldDisplayOrder(titles).build(list);
        FileExportUtil.export(workbook, new File(TEST_DIR + "map_build.xlsx"));
    }

    @Test
    void commonBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hideColumns(0, 1)
                .globalStyle("background-color:red;", "title->background-color:yellow;")
                .start()) {
            data(excelBuilder, 100);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "common_build.xlsx"));
        }
    }

    @Test
    void hasStyleBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
                .start()) {
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "has_style_build.xlsx"));
        }
    }

    @Test
    void customWidthBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .widths(15, 20, 25, 30)
                .start()) {
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "custom_width_build.xlsx"));
        }
    }

    @Test
    void continueBuild() throws Exception {
        DefaultStreamExcelBuilder<CommonPeople> excelBuilder = null;
        try {
            excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                    .fixedTitles()
                    .widths(15, 20, 25, 30)
                    .start();
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();

            excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class, workbook)
                    .fixedTitles()
                    .start();
            data(excelBuilder, 10000);
            FileExportUtil.export(workbook, new File(TEST_DIR + "continue_build.xlsx"));
        } catch (Throwable e) {
            if (excelBuilder != null) {
                excelBuilder.clear();
            }
            throw new RuntimeException(e);
        }
    }

    @Test
    void cancelBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .widths(15, 20, 25, 30)
                .start()) {
            data(excelBuilder, 10000);
            excelBuilder.cancel();
        }
    }

    @Test
    void buildAsPaths() throws Exception {
        List<Path> paths = null;
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .widths(15, 20, 25, 30)
                .capacity(1000)
                .start()) {
            data(excelBuilder, 10000);
            paths = excelBuilder.buildAsPaths();
        } finally {
            TempFileOperator.deleteTempFiles(paths);
        }
    }

    @Test
    void buildAsZip() throws Exception {
        Path zip = null;
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .widths(15, 20, 25, 30)
                .capacity(1000)
                .start()) {
            data(excelBuilder, 10000);
            zip = excelBuilder.buildAsZip("test");
        } finally {
            TempFileOperator.deleteTempFile(zip);
        }
    }

    @Test
    void bigBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .widths(15, 20, 25, 30)
                .start()) {
            data(excelBuilder, 1200000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "big_build.xlsx"));
        }
    }

    @Test
    void customStyleBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CustomStylePeople> excelBuilder = DefaultStreamExcelBuilder.of(CustomStylePeople.class)
                .fixedTitles()
                .hasStyle()
                .start()) {
            customStyleData(excelBuilder, 1000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "custom_style_build.xlsx"));
        }
    }

    @Test
    void evenOddBuild() throws Exception {
        try (DefaultStreamExcelBuilder<OddEvenStylePeople> excelBuilder = DefaultStreamExcelBuilder.of(OddEvenStylePeople.class)
                .fixedTitles()
                .hasStyle()
                .start()) {
            oddEvenData(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "odd_even_build.xlsx"));
        }
    }

    @Test
    void groupBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .groups(CommonPeople.class)
                .widths(50)
                .start()) {
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "group_build.xlsx"));
        }
    }

    @Test
    void widthBuild() throws Exception {
        try (DefaultStreamExcelBuilder<WidthPeople> excelBuilder = DefaultStreamExcelBuilder.of(WidthPeople.class)
                .fixedTitles()
                .start()) {
            widthEvenData(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "width_build.xlsx"));
        }
    }

    @Test
    void formulaBuild() throws Exception {
        try (DefaultStreamExcelBuilder<Formula> excelBuilder = DefaultStreamExcelBuilder.of(Formula.class)
                .fixedTitles()
                .start()) {
            formulaData(excelBuilder, 100);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "formula_build.xlsx"));
        }
    }

    private void data(DefaultStreamExcelBuilder<CommonPeople> excelBuilder, int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        for (int i = 0; i < size; i++) {
            CommonPeople commonPeople = new CommonPeople();
            boolean odd = i % 2 == 0;
            commonPeople.setName(odd ? "张三" : "李四");
            commonPeople.setAge(odd ? 18 : 24);
            commonPeople.setDance(odd ? true : false);
            commonPeople.setMoney(odd ? oddMoney : evenMoney);
            commonPeople.setBirthday(new Date());
            commonPeople.setLocalDate(LocalDate.now());
            commonPeople.setLocalDateTime(LocalDateTime.now());
            commonPeople.setCats(100L);
            excelBuilder.append(commonPeople);
        }
    }

    private void customStyleData(DefaultStreamExcelBuilder<CustomStylePeople> excelBuilder, int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        for (int i = 0; i < size; i++) {
            CustomStylePeople customStylePeople = new CustomStylePeople();
            boolean odd = i % 2 == 0;
            customStylePeople.setName(odd ? "张三" : "李四");
            customStylePeople.setAge(odd ? 18 : 24);
            customStylePeople.setDance(odd ? true : false);
            customStylePeople.setMoney(odd ? oddMoney : evenMoney);
            excelBuilder.append(customStylePeople);
        }
    }

    private void oddEvenData(DefaultStreamExcelBuilder<OddEvenStylePeople> excelBuilder, int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        for (int i = 0; i < size; i++) {
            OddEvenStylePeople oddEvenStylePeople = new OddEvenStylePeople();
            boolean odd = i % 2 == 0;
            oddEvenStylePeople.setName(odd ? "张三" : "李四");
            oddEvenStylePeople.setAge(odd ? 18 : 24);
            oddEvenStylePeople.setDance(odd ? true : false);
            oddEvenStylePeople.setMoney(odd ? oddMoney : evenMoney);
            excelBuilder.append(oddEvenStylePeople);
        }
    }

    private void widthEvenData(DefaultStreamExcelBuilder<WidthPeople> excelBuilder, int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        for (int i = 0; i < size; i++) {
            WidthPeople oddEvenStylePeople = new WidthPeople();
            boolean odd = i % 2 == 0;
            oddEvenStylePeople.setName(odd ? "张三" : "李四");
            oddEvenStylePeople.setAge(odd ? 18 : 24);
            oddEvenStylePeople.setDance(odd ? true : false);
            oddEvenStylePeople.setMoney(odd ? oddMoney : evenMoney);
            excelBuilder.append(oddEvenStylePeople);
        }
    }

    private void formulaData(DefaultStreamExcelBuilder<Formula> excelBuilder, int size) {
        for (int i = 0; i < 100; i++) {
            Formula formula = new Formula();
            excelBuilder.append(formula);
        }
    }
}