/*
 * Copyright 2017 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.html2excel.core;

import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultExcelBuilder {

    private static final String DEFAULT_TEMPLATE_PATH = "/template/beetl/defaultExcelBuilderTemplate.html";

    private ExcelBuilder excelBuilder;

    private Map<String, String> titleFieldMapping;

    private DefaultExcelBuilder() {
        this.excelBuilder = new BeetlExcelBuilder();
    }

    public static DefaultExcelBuilder getInstance() {
        return new DefaultExcelBuilder();
    }

    public DefaultExcelBuilder mapping(LinkedHashMap<String, String> titleFieldMapping) {
        if (Objects.isNull(titleFieldMapping) || titleFieldMapping.isEmpty()) {
            throw new IllegalArgumentException("TitleFieldMapping is necessary");
        }
        this.titleFieldMapping = titleFieldMapping;
        return this;
    }

    public Workbook build(List<?> data) {
        if (Objects.isNull(titleFieldMapping) || titleFieldMapping.isEmpty()) {
            throw new IllegalArgumentException("TitleFieldMapping is necessary");
        }
        List<String> title = new ArrayList<>(titleFieldMapping.size());
        titleFieldMapping.forEach((k, v) -> title.add(k));

        Map<String, Object> renderData = new HashMap<>();
        renderData.put("title", title);

        if (Objects.isNull(data) || data.isEmpty()) {
            return excelBuilder.useDefaultStyle().build(renderData);
        }
        Class<?> clazz = data.get(0).getClass();
        Field[] fields = clazz.getDeclaredFields();
        if (Objects.isNull(fields) || fields.length == 0) {
            return excelBuilder.useDefaultStyle().build(renderData);
        }

        Map<String, Field> fieldMap = Arrays.stream(fields)
                .peek(field -> field.setAccessible(true))
                .filter(field -> Modifier.isPrivate(field.getModifiers()))
                .collect(Collectors.toMap(Field::getName, f -> f));

        List<Field> sortedField = new ArrayList<>(fields.length);
        titleFieldMapping.forEach((k, v) -> sortedField.add(fieldMap.get(v)));

        List<List<Object>> contents = data.stream().map(d ->
                sortedField.stream().map(f -> {
                    try {
                        return f.get(d);
                    } catch (Exception e) {
                        throw new RuntimeException(e);
                    }
                }).collect(Collectors.toList())).collect(Collectors.toList());
        renderData.put("contents", contents);
        return excelBuilder.template(DEFAULT_TEMPLATE_PATH).useDefaultStyle().build(renderData);
    }
}
