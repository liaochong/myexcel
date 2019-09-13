package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.pojo.CsvPeople;
import com.github.liaochong.myexcel.utils.FileExportUtil;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Paths;
import java.util.LinkedList;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
class CsvBuilderTest extends BasicTest {

    @Test
    void build() {
        Csv csv = CsvBuilder.of(CsvPeople.class).build(data(10000));
        FileExportUtil.export(csv.getFilePath(), Paths.get(TEST_DIR + "common.csv"));
    }

    @Test
    void appendBuild() {
        CsvBuilder<CsvPeople> csvBuilder = CsvBuilder.of(CsvPeople.class);
        for (int i = 0; i < 10; i++) {
            csvBuilder.append(data(1000));
        }
        Csv csv = csvBuilder.build();
        FileExportUtil.export(csv.getFilePath(), Paths.get(TEST_DIR + "append.csv"));
    }

    private List<CsvPeople> data(int size) {
        BigDecimal oddMoney = new BigDecimal(109898);
        BigDecimal evenMoney = new BigDecimal(66666);

        List<CsvPeople> peoples = new LinkedList<>();
        for (int i = 0; i < size; i++) {
            CsvPeople csvPeople = new CsvPeople();
            boolean odd = i % 2 == 0;
            csvPeople.setName(odd ? "张三" : "李四");
            csvPeople.setAge(odd ? 18 : 24);
            csvPeople.setDance(odd ? true : false);
            csvPeople.setMoney(odd ? oddMoney : evenMoney);
            peoples.add(csvPeople);
        }
        return peoples;
    }

}