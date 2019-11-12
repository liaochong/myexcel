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

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.converter.ReadConverterContext;
import com.github.liaochong.myexcel.exception.StopReadException;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.extern.slf4j.Slf4j;

import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

/**
 * @author liaochong
 * @version 1.0
 */
@Slf4j
class CsvReadHandler<T> extends AbstractReadHandler<T> {

    private static final Pattern PATTERN_SPLIT = Pattern.compile(",(?=([^\\\"]*\\\"[^\\\"]*\\\")*[^\\\"]*$)");

    private static final Pattern PATTERN_QUOTES = Pattern.compile("[\"]{2}");

    private InputStream is;

    private String charset;

    public CsvReadHandler(InputStream is,
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
        this.charset = readConfig.getCharset();
        this.exceptionFunction = readConfig.getExceptionFunction();
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
                this.initFieldMap(lineIndex);
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

    @SuppressWarnings("unchecked")
    private void process(String line, Row row) {
        if (!rowFilter.test(row)) {
            return;
        }
        T obj = this.newInstance(dataType);
        if (line != null) {
            String[] strArr = PATTERN_SPLIT.split(line, -1);
            for (int i = 0, size = strArr.length; i < size; i++) {
                String content = strArr[i];
                if (content != null && content.isEmpty()) {
                    content = null;
                }
                if (content != null && content.indexOf(Constants.QUOTES) == 0) {
                    if (content.length() > 2) {
                        content = content.substring(1, content.length() - 1);
                    } else {
                        content = "";
                    }
                }
                if (content != null) {
                    content = PATTERN_QUOTES.matcher(content).replaceAll("\"");
                }
                this.addTitles(content, row.getRowNum(), i);
                if (isMapType) {
                    ((Map<Cell, String>) obj).put(new Cell(row.getRowNum(), i), content);
                    continue;
                }
                Field field = fieldMap.get(i);
                if (field == null) {
                    continue;
                }
                ReadContext<T> context = new ReadContext<>(obj, field, content, row.getRowNum(), i);
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
