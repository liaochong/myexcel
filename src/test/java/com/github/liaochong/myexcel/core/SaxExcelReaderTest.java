package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CommonPeople;
import com.github.liaochong.myexcel.core.pojo.CsvPeople;
import com.github.liaochong.myexcel.core.pojo.ExceptionPeople;
import org.junit.jupiter.api.Test;

import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
class SaxExcelReaderTest {

    @Test
    void csvReadFile() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common.csv");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<CsvPeople> csvPeoples = SaxExcelReader.of(CsvPeople.class).rowFilter(row -> row.getRowNum() > 0).read(path.toFile());
    }

    @Test
    void csvReadContinuedException() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common.csv");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<ExceptionPeople> csvPeoples = SaxExcelReader.of(ExceptionPeople.class).exceptionally((e, context) -> {
            System.out.println(context.getField().getName() + "_" + context.getVal() + "_" + context.getRowNum() + "_" + context.getColNum());
            return true;
        }).read(path.toFile());
    }

    @Test
    void csvReadException() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common.csv");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        try {
            List<ExceptionPeople> csvPeoples = SaxExcelReader.of(ExceptionPeople.class).exceptionally((e, context) -> {
                System.out.println(context.getField().getName() + "_" + context.getVal() + "_" + context.getRowNum() + "_" + context.getColNum());
                return false;
            }).read(path.toFile());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void readXlsxFile() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<CommonPeople> commonPeoples = SaxExcelReader.of(CommonPeople.class).rowFilter(row -> row.getRowNum() > 0).read(path.toFile());
    }

    @Test
    void readXlsFile() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xls");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<CommonPeople> commonPeoples = SaxExcelReader.of(CommonPeople.class).rowFilter(row -> row.getRowNum() > 0).read(path.toFile());
    }

    @Test
    void readXlsxContinuedException() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<CommonPeople> commonPeoples = SaxExcelReader.of(CommonPeople.class).exceptionally((e, context) -> {
            System.out.println(context.getField().getName() + "_" + context.getVal() + "_" + context.getRowNum() + "_" + context.getColNum());
            return true;
        }).read(path.toFile());
    }

    @Test
    void readXlsxException() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        try {
            List<CommonPeople> commonPeoples = SaxExcelReader.of(CommonPeople.class).exceptionally((e, context) -> {
                System.out.println(context.getField().getName() + "_" + context.getVal() + "_" + context.getRowNum() + "_" + context.getColNum());
                return false;
            }).read(path.toFile());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void readXlsContinuedException() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xls");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<ExceptionPeople> commonPeoples = SaxExcelReader.of(ExceptionPeople.class).exceptionally((e, context) -> {
            System.out.println(context.getField().getName() + "_" + context.getVal() + "_" + context.getRowNum() + "_" + context.getColNum());
            return true;
        }).read(path.toFile());
    }

    @Test
    void readXlsException() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xls");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        try {
            List<ExceptionPeople> commonPeoples = SaxExcelReader.of(ExceptionPeople.class).exceptionally((e, context) -> {
                System.out.println(context.getField().getName() + "_" + context.getVal() + "_" + context.getRowNum() + "_" + context.getColNum());
                return false;
            }).read(path.toFile());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void readThen() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        SaxExcelReader.of(CsvPeople.class).rowFilter(row -> row.getRowNum() > 0).readThen(path.toFile(), d -> {
            System.out.println(d.getMoney());
        });
    }

    @Test
    void readThenFile() {
    }

    @Test
    void readThenInputStream() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        SaxExcelReader.of(CsvPeople.class).rowFilter(row -> row.getRowNum() > 0).readThen(Files.newInputStream(path), d -> {
            System.out.println(d.getMoney());
        });
    }

    @Test
    void readThen3() {
    }
}