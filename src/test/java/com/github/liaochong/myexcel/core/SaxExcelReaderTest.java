package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CommonPeople;
import com.github.liaochong.myexcel.core.pojo.CsvPeople;
import com.sun.tools.javac.util.Assert;
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
        Assert.check(csvPeoples.size() == 1000);
    }

    @Test
    void readXlsxFile() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<CommonPeople> commonPeoples = SaxExcelReader.of(CommonPeople.class).rowFilter(row -> row.getRowNum() > 0).read(path.toFile());
        Assert.check(commonPeoples.size() == 10000);
    }

    @Test
    void readXlsFile() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xls");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<CommonPeople> commonPeoples = SaxExcelReader.of(CommonPeople.class).rowFilter(row -> row.getRowNum() > 0).read(path.toFile());
        Assert.check(commonPeoples.size() == 10000);
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