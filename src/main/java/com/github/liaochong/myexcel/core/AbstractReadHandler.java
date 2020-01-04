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
import com.github.liaochong.myexcel.utils.ReflectUtil;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;
import java.util.function.Predicate;

/**
 * 读取抽象
 *
 * @author liaochong
 * @version 1.0
 */
abstract class AbstractReadHandler<T> {

    protected boolean isMapType;

    protected Map<Integer, Field> fieldMap;

    protected Class<T> dataType;

    protected Consumer<T> consumer;

    protected Function<T, Boolean> function;

    protected Predicate<Row> rowFilter;

    protected Predicate<T> beanFilter;

    protected List<T> result;

    protected T obj;

    protected Map<String, Integer> titles = new HashMap<>();

    protected BiFunction<Throwable, ReadContext, Boolean> exceptionFunction;

    protected SaxExcelReader.ReadConfig<T> readConfig;

    protected AddTitleConsumer<String, Integer, Integer> addTitleConsumer = (v, rowNum, colNum) -> {
    };

    private ReadContext<T> context = new ReadContext<>();

    protected void init(
            List<T> result,
            SaxExcelReader.ReadConfig<T> readConfig) {
        this.result = result;
        dataType = readConfig.getDataType();
        fieldMap = ReflectUtil.getFieldMapOfExcelColumn(dataType);
        consumer = readConfig.getConsumer();
        function = readConfig.getFunction();
        rowFilter = readConfig.getRowFilter();
        beanFilter = readConfig.getBeanFilter();
        exceptionFunction = readConfig.getExceptionFunction();
        this.readConfig = readConfig;
        if (fieldMap.isEmpty()) {
            addTitleConsumer = this::addTitles;
        }
    }

    @SuppressWarnings("unchecked")
    T newInstance(Class<T> clazz) {
        if (isMapType) {
            return (T) new LinkedHashMap<Cell, String>();
        }
        if (clazz == Map.class) {
            isMapType = true;
            return (T) new LinkedHashMap<Cell, String>();
        }
        try {
            return clazz.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    protected void initFieldMap(int rowNum) {
        if (rowNum != 0 || !fieldMap.isEmpty()) {
            return;
        }
        Map<String, Field> titleFieldMap = ReflectUtil.getFieldMapOfTitleExcelColumn(dataType);
        fieldMap = new HashMap<>(titleFieldMap.size());
        titles.forEach((k, v) -> {
            fieldMap.put(v, titleFieldMap.get(k));
        });
    }

    protected void convert(String value, int rowNum, int colNum, Field field) {
        if (value == null || field == null) {
            return;
        }
        context.reset(obj, field, value, rowNum, colNum);
        ReadConverterContext.convert(obj, context, exceptionFunction);
    }

    private void addTitles(String formattedValue, int rowNum, int thisCol) {
        if (rowNum == 0) {
            titles.put(formattedValue, thisCol);
        }
    }
}
