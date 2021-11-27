/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core;

import java.io.File;
import java.io.InputStream;
import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * 字符列sax读取，原理是使用Map接受值，选择指定的某一列
 *
 * @author liaochong
 * @version 1.0
 */
public class StringColumnSaxExcelReader {

    private SaxExcelReader<Map> saxExcelReader = SaxExcelReader.of(Map.class);
    /**
     * 默认取第一列
     */
    private final int columnNum;

    private StringColumnSaxExcelReader(int columnNum) {
        this.columnNum = columnNum;
    }

    public static StringColumnSaxExcelReader columnNum(int columnNum) {
        return new StringColumnSaxExcelReader(columnNum);
    }

    public StringColumnSaxExcelReader rowFilter(Predicate<Row> predicate) {
        saxExcelReader.rowFilter(predicate);
        return this;
    }

    public StringColumnSaxExcelReader sheet(int sheetNum) {
        saxExcelReader.sheet(sheetNum);
        return this;
    }

    public StringColumnSaxExcelReader sheet(String sheetName) {
        saxExcelReader.sheet(sheetName);
        return this;
    }

    public List<String> read(InputStream inputStream) {
        List<Map> result = saxExcelReader.read(inputStream);
        return mapToString(result);
    }

    public List<String> read(File file) {
        List<Map> result = saxExcelReader.read(file);
        return mapToString(result);
    }

    private List<String> mapToString(List<Map> result) {
        if (result == null || result.isEmpty()) {
            return Collections.emptyList();
        }
        return result.stream().map(map -> ((Set<Cell>) map.keySet()).stream().filter(cell -> cell.getColNum() == columnNum)
                        .map(((Map<Cell, String>) map)::get)
                        .findFirst().orElse(null))
                .collect(Collectors.toCollection(LinkedList::new));
    }
}
