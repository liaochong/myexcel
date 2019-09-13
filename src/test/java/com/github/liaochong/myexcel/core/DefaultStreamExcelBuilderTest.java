package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CommonPeople;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.math.BigDecimal;

/**
 * @author liaochong
 * @version 1.0
 */
class DefaultStreamExcelBuilderTest extends BasicTest {

    @Test
    void commonBuild() throws Exception {
        try (DefaultStreamExcelBuilder excelBuilder = DefaultStreamExcelBuilder.of(CommonPeople.class)
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
                    .workbookType(WorkbookType.XLSX)
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
    void buildAsPaths() {
    }

    @Test
    void buildAsZip() {
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
}