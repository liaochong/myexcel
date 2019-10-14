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
import lombok.extern.slf4j.Slf4j;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.reflect.Field;
import java.nio.file.Files;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class CsvHandler<T> {

    private final Map<Integer, Field> fieldMap;

    private InputStream is;

    private List<T> result;

    private Class<T> dataType;

    private Consumer<T> consumer;

    private Function<T, Boolean> function;

    private Predicate<Row> rowFilter;

    private Predicate<T> beanFilter;

    private String charset;

    private BiFunction<Throwable, ReadContext, Boolean> exceptionFunction;

    public CsvHandler(InputStream is,
                      SaxExcelReader.ReadConfig<T> readConfig,
                      List<T> result) {
        this.is = is;
        try {
            is.reset();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        this.result = result;
        this.dataType = readConfig.getDataType();
        this.fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        this.consumer = readConfig.getConsumer();
        this.function = readConfig.getFunction();
        this.rowFilter = readConfig.getRowFilter();
        this.beanFilter = readConfig.getBeanFilter();
        this.charset = readConfig.getCharset();
        this.exceptionFunction = readConfig.getExceptionFunction();
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
            this.charset = readConfig.getCharset();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void read() {
        if (is == null) {
            return;
        }
        long startTime = System.currentTimeMillis();
        try (BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(is, charset))) {
            int lineIndex = 0;
            String line;
            while ((line = bufferedReader.readLine()) != null) {
                Row row = new Row(lineIndex);
                this.process(line, row);
                lineIndex++;
            }
            log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
        } catch (StopReadException e) {
            log.info("Sax import takes {} ms", System.currentTimeMillis() - startTime);
            throw e;
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
            String[] strArr = line.split(",(?=([^\\\"]*\\\"[^\\\"]*\\\")*[^\\\"]*$)", -1);
            for (int i = 0, size = strArr.length; i < size; i++) {
                String content = strArr[i];
                Field field = fieldMap.get(i);
                if (field == null) {
                    continue;
                }
                ReadContext context = new ReadContext(field, content, row.getRowNum(), i);
                ReadConverterContext.convert(obj, context, exceptionFunction);
            }
        }
        if (!beanFilter.test(obj)) {
            return;
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
