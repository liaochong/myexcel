package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CommonPeople;
import com.github.liaochong.myexcel.core.pojo.CustomStylePeople;
import com.github.liaochong.myexcel.core.pojo.OddEvenStylePeople;
import com.github.liaochong.myexcel.core.pojo.WidthPeople;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import com.github.liaochong.myexcel.utils.TempFileOperator;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
class DefaultStreamExcelBuilderTest extends BasicTest {

    @Test
    void commonBuild() throws Exception {
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .workbookType(WorkbookType.XLS)
                .fixedTitles()
                .start()) {
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "common_build.xlsx"));
        }
    }

    @Test
    void hasStyleBuild() throws Exception {
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
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
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
                .widths(15, 20, 25, 30)
                .start()) {
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "custom_width_build.xlsx"));
        }
    }

    @Test
    void continueBuild() throws Exception {
        DefaultStreamExcelBuilder excelBuilder = null;
        try {
            excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                    .fixedTitles()
                    .hasStyle()
                    .widths(15, 20, 25, 30)
                    .start();
            data(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();

            excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class, workbook)
                    .fixedTitles()
                    .hasStyle()
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
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
                .widths(15, 20, 25, 30)
                .start()) {
            data(excelBuilder, 10000);
            excelBuilder.cancle();
        }
    }

    @Test
    void buildAsPaths() throws Exception {
        List<Path> paths = null;
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
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
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
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
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
                .widths(15, 20, 25, 30)
                .start()) {
            data(excelBuilder, 1200000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "big_build.xlsx"));
        }
    }

    @Test
    void customStyleBuild() throws Exception {
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CustomStylePeople.class)
                .fixedTitles()
                .hasStyle()
                .start()) {
            customStyleData(excelBuilder, 1200000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "custom_style_build.xlsx"));
        }
    }

    @Test
    void evenOddBuild() throws Exception {
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(OddEvenStylePeople.class)
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
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
                .fixedTitles()
                .hasStyle()
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
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(WidthPeople.class)
                .fixedTitles()
                .hasStyle()
                .start()) {
            widthEvenData(excelBuilder, 10000);
            Workbook workbook = excelBuilder.build();
            FileExportUtil.export(workbook, new File(TEST_DIR + "width_build.xlsx"));
        }
    }

    private void data(DefaultStreamExcelBuilder excelBuilder, int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);
        for (int i = 0; i < size; i++) {
            CommonPeople commonPeople = new CommonPeople();
            boolean odd = i % 2 == 0;
            commonPeople.setName(odd ? "张三" : "李四");
            commonPeople.setAge(odd ? 18 : 24);
            commonPeople.setDance(odd ? true : false);
            commonPeople.setMoney(odd ? oddMoney : evenMoney);
            excelBuilder.append(commonPeople);
        }
    }

    private void customStyleData(DefaultStreamExcelBuilder excelBuilder, int size) {
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

    private void oddEvenData(DefaultStreamExcelBuilder excelBuilder, int size) {
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

    private void widthEvenData(DefaultStreamExcelBuilder excelBuilder, int size) {
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
}