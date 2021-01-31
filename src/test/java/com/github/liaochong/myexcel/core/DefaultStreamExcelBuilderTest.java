package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CommonPeople;
import com.github.liaochong.myexcel.core.pojo.CustomStylePeople;
import com.github.liaochong.myexcel.core.pojo.Formula;
import com.github.liaochong.myexcel.core.pojo.MultiPeople;
import com.github.liaochong.myexcel.core.pojo.OddEvenStylePeople;
import com.github.liaochong.myexcel.core.pojo.Product;
import com.github.liaochong.myexcel.core.pojo.WidthPeople;
import com.github.liaochong.myexcel.core.templatehandler.FreemarkerTemplateHandler;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.Executors;

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

        Workbook workbook = DefaultExcelBuilder.of(Map.class)
                .style("cell&1->background-color:red;", "title&2->color:green;font-weight:bold;background-color:yellow;")
                .width(1, 20)
                .fieldDisplayOrder(titles).titles(titles).build(list);
//        workbook = DefaultExcelBuilder.of(Map.class).fieldDisplayOrder(titles).build(list);
        FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "map_build.xlsx"));
    }

    @Test
    void commonBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
//                .hideColumns(0, 1)
//                .globalStyle("background-color:red;", "title->background-color:yellow;")
                .start()) {
            data(excelBuilder, 5000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "common_build.xlsx"));
        }
    }

    @Test
    void asyncCommonBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .threadPool(Executors.newFixedThreadPool(10))
                .start()) {
            for (int i = 0; i < 1000; i++) {
                excelBuilder.asyncAppend(this::dataList);
            }
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "async_common_build.xlsx"));
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
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "custom_width_build.xlsx"));
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
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "continue_build.xlsx"));
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
                .start()) {
            data(excelBuilder, 65500);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "big_build.xlsx"));
        }
    }

    @Test
    void customStyleBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CustomStylePeople> excelBuilder = DefaultStreamExcelBuilder.of(CustomStylePeople.class)
                .fixedTitles()
                .start()) {
            customStyleData(excelBuilder, 1000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "custom_style_build.xlsx"));
        }
    }

    @Test
    void evenOddBuild() throws Exception {
        try (DefaultStreamExcelBuilder<OddEvenStylePeople> excelBuilder = DefaultStreamExcelBuilder.of(OddEvenStylePeople.class)
                .fixedTitles()
                .style("cell&1->color:green;")
                .start()) {
            oddEvenData(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "odd_even_build.xlsx"));
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
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "group_build.xlsx"));
        }
    }

    @Test
    void widthBuild() throws Exception {
        try (DefaultStreamExcelBuilder<WidthPeople> excelBuilder = DefaultStreamExcelBuilder.of(WidthPeople.class)
                .fixedTitles()
                .start()) {
            widthEvenData(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "width_build.xlsx"));
        }
    }

    @Test
    void formulaBuild() throws Exception {
        try (DefaultStreamExcelBuilder<Formula> excelBuilder = DefaultStreamExcelBuilder.of(Formula.class)
                .fixedTitles()
                .start()) {
            formulaData(excelBuilder, 100);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "formula_build.xlsx"));
        }
    }

    @Test
    void appendTemplateBuild() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .templateHandler(FreemarkerTemplateHandler.class)
                .start()) {
            data(excelBuilder, 100);
            excelBuilder.append("/templates/freemarkerToExcelExample.ftl", getDataMap());
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "append_template_build.xlsx"));
        }
    }

    @Test
    void appendExistExcel() throws Exception {
        try (DefaultStreamExcelBuilder<CommonPeople> excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class, Paths.get(TEST_OUTPUT_DIR + "common_build.xlsx"))
                .fixedTitles()
                .start()) {
            data(excelBuilder, 5000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "common_build.xlsx"));
        }
    }

    @Test
    void testMultiColumn() throws Exception {
        try (DefaultStreamExcelBuilder<MultiPeople> excelBuilder = DefaultStreamExcelBuilder.of(MultiPeople.class)
                .fixedTitles()
                .freezePane(new FreezePane(2, 5))
                .start()) {
            multiData(excelBuilder);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_OUTPUT_DIR + "multi_build.xlsx"));
        }
    }

    private void multiData(DefaultStreamExcelBuilder<MultiPeople> excelBuilder) {
        MultiPeople people = new MultiPeople();
        people.setTastes(new LinkedList<>());
        people.setDates(new LinkedList<>());
        people.setName("姓名");
        for (int i = 0; i < 5; i++) {
            people.getTastes().add("兴趣" + i);
        }
        for (int i = 0; i < 10; i++) {
            people.getDates().add(new Date());
        }
        excelBuilder.append(people);

    }

    private void data(DefaultStreamExcelBuilder<CommonPeople> excelBuilder, int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        List<CompletableFuture> futures = new LinkedList<>();
        for (int i = 0; i < size; i++) {
            int index = i;
            CompletableFuture future = CompletableFuture.runAsync(() -> {
                CommonPeople commonPeople = new CommonPeople();
                boolean odd = index % 2 == 0;
                commonPeople.setName(odd ? "张三" : "李四");
                commonPeople.setAge(odd ? 18 : 24);
                commonPeople.setDance(odd ? true : false);
                commonPeople.setMoney(odd ? oddMoney : evenMoney);
                commonPeople.setBirthday(new Date());
                commonPeople.setLocalDate(LocalDate.now());
                commonPeople.setLocalDateTime(LocalDateTime.now());
                commonPeople.setCats(100L);
                excelBuilder.append(commonPeople);
            });
            futures.add(future);
        }
        futures.forEach(CompletableFuture::join);
    }

    private List<CommonPeople> dataList() {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        List<CommonPeople> result = new ArrayList<>();
        for (int i = 0; i < 1000; i++) {
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
            result.add(commonPeople);
        }
        return result;
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