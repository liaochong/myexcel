package com.github.liaochong.myexcel.core;

import org.junit.jupiter.api.Test;

import java.net.URL;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
class StringColumnSaxExcelReaderTest {

    @Test
    public void testRead() throws Exception {
        URL htmlToExcelEampleURL = this.getClass().getResource("/common_build.xlsx");
        Path path = Paths.get(htmlToExcelEampleURL.toURI());

        List<String> strings = StringColumnSaxExcelReader.columnNum(0)
                .rowFilter(row -> row.getRowNum() > 1)
                .read(path.toFile());
        System.out.println(strings.size());
    }

}