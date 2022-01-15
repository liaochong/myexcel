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
import java.util.Objects;
import java.util.Set;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.stream.Collectors;

/**
 * 字符列sax读取，原理是使用Map接受值，选择指定的某一列
 *
 * @author liaochong
 * @version 4.0.0.RC
 */
public class ColumnSaxExcelReader {

    private final SaxExcelReader<Map> saxExcelReader = SaxExcelReader.of(Map.class);
    /**
     * 默认取第一列
     */
    private final int columnNum;
    /**
     * 是否忽略空单元格
     */
    private boolean ignoreBlankCell;

    private ColumnSaxExcelReader(int columnNum) {
        this.columnNum = columnNum;
    }

    public static ColumnSaxExcelReader columnNum(int columnNum) {
        return new ColumnSaxExcelReader(columnNum);
    }

    public ColumnSaxExcelReader rowFilter(Predicate<Row> predicate) {
        saxExcelReader.rowFilter(predicate);
        return this;
    }

    public ColumnSaxExcelReader sheet(int sheetNum) {
        saxExcelReader.sheet(sheetNum);
        return this;
    }

    public ColumnSaxExcelReader sheet(String sheetName) {
        saxExcelReader.sheet(sheetName);
        return this;
    }

    public ColumnSaxExcelReader ignoreBlankCell() {
        saxExcelReader.ignoreBlankRow();
        ignoreBlankCell = true;
        return this;
    }

    public ColumnSaxExcelReader stopReadingOnBlankRow() {
        saxExcelReader.stopReadingOnBlankRow();
        return this;
    }

    public List<String> readAsString(InputStream inputStream) {
        List<Map> result = saxExcelReader.read(inputStream);
        return mapToString(result);
    }

    public List<String> readAsString(File file) {
        List<Map> result = saxExcelReader.read(file);
        return mapToString(result);
    }

    public List<Integer> readAsInteger(InputStream inputStream) {
        return this.getList(saxExcelReader.read(inputStream), Integer::valueOf);
    }

    public List<Integer> readAsInteger(File file) {
        return this.getList(saxExcelReader.read(file), Integer::valueOf);
    }

    public List<Long> readAsLong(InputStream inputStream) {
        return getList(saxExcelReader.read(inputStream), Long::valueOf);
    }

    public List<Long> readAsLong(File file) {
        return getList(saxExcelReader.read(file), Long::valueOf);
    }

    public List<Boolean> readAsBoolean(InputStream inputStream) {
        return getList(saxExcelReader.read(inputStream), Boolean::valueOf);
    }

    public List<Boolean> readAsBoolean(File file) {
        return getList(saxExcelReader.read(file), Boolean::valueOf);
    }

    public List<Double> readAsDouble(InputStream inputStream) {
        return getList(saxExcelReader.read(inputStream), Double::valueOf);
    }

    public List<Double> readAsDouble(File file) {
        return getList(saxExcelReader.read(file), Double::valueOf);
    }

    public List<Short> readAsShort(InputStream inputStream) {
        return getList(saxExcelReader.read(inputStream), Short::valueOf);
    }

    public List<Short> readAsShort(File file) {
        return getList(saxExcelReader.read(file), Short::valueOf);
    }

    public List<Float> readAsFloat(InputStream inputStream) {
        return getList(saxExcelReader.read(inputStream), Float::valueOf);
    }

    public List<Float> readAsFloat(File file) {
        return getList(saxExcelReader.read(file), Float::valueOf);
    }

    public List<Byte> readAsByte(InputStream inputStream) {
        return getList(saxExcelReader.read(inputStream), Byte::valueOf);
    }

    public List<Byte> readAsByte(File file) {
        return getList(saxExcelReader.read(file), Byte::valueOf);
    }

    @SuppressWarnings("unchecked")
    private List<String> mapToString(List<Map> result) {
        if (result == null || result.isEmpty()) {
            return Collections.emptyList();
        }
        return result.stream().map(map -> ((Set<Cell>) map.keySet()).stream().filter(cell -> cell.getColNum() == columnNum)
                        .map(((Map<Cell, String>) map)::get)
                        .filter(Objects::nonNull)
                        .findFirst().orElse(null))
                .filter(v -> !ignoreBlankCell || v != null)
                .collect(Collectors.toCollection(LinkedList::new));
    }

    private <T> List<T> getList(List<Map> result, Function<String, T> function) {
        List<String> strings = mapToString(result);
        return strings.stream().map(s -> s == null ? null : function.apply(s)).collect(Collectors.toCollection(LinkedList::new));
    }
}
