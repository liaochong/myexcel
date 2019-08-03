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

import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import com.github.liaochong.myexcel.exception.StopReadException;
import com.github.liaochong.myexcel.utils.ReflectUtil;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * @author liaochong
 * @version 1.0
 */
class CsvHandler<T> {

    private final Map<Integer, Field> fieldMap;

    private InputStream is;

    private List<T> result;

    private Class<T> dataType;

    private Consumer<T> consumer;

    private Function<T, Boolean> function;

    private Predicate<Row> rowFilter;

    private Predicate<T> beanFilter;

    public CsvHandler(InputStream is,
                      SaxExcelReader.ReadConfig<T> readConfig,
                      List<T> result) {
        this.is = is;
        this.result = result;
        this.dataType = readConfig.getDataType();
        this.fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        this.consumer = readConfig.getConsumer();
        this.function = readConfig.getFunction();
        this.rowFilter = readConfig.getRowFilter();
        this.beanFilter = readConfig.getBeanFilter();
    }

    public CsvHandler(File file,
                      SaxExcelReader.ReadConfig<T> readConfig,
                      List<T> result) {
        try {
            this.is = Files.newInputStream(file.toPath());
            this.result = result;
            this.dataType = readConfig.getDataType();
            this.fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
            this.consumer = readConfig.getConsumer();
            this.function = readConfig.getFunction();
            this.rowFilter = readConfig.getRowFilter();
            this.beanFilter = readConfig.getBeanFilter();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void read() {
        if (is == null) {
            return;
        }
        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(is));
        try {
            int lineIndex = 0;
            String line = bufferedReader.readLine();
            while (line != null) {
                Row row = new Row(lineIndex);
                this.process(line, row);
                line = bufferedReader.readLine();
                lineIndex++;
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void process(String line, Row row) throws Exception {
        if (!rowFilter.test(row)) {
            return;
        }
        T obj = dataType.newInstance();
        if (line != null) {
            String[] strArr = line.trim().split(",(?=([^\\\"]*\\\"[^\\\"]*\\\")*[^\\\"]*$)", -1);
            for (int i = 0, size = strArr.length; i < size; i++) {
                String content = strArr[i];
                Field field = fieldMap.get(i);
                if (Objects.isNull(field)) {
                    continue;
                }
                ReadConverterContext.convert(content, field, obj);
            }
        }
        if (consumer != null) {
            consumer.accept(obj);
        } else if (function != null) {
            Boolean noStop = function.apply(obj);
            if (!noStop) {
                throw new StopReadException();
            }
        } else {
            result.add(obj);
        }
    }
}
