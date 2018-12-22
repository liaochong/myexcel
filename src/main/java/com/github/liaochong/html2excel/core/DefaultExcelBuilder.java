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

import com.github.liaochong.html2excel.core.reflect.ClassFieldContainer;
import com.github.liaochong.html2excel.utils.ReflectUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelBuilder {

    private static final String DEFAULT_TEMPLATE_PATH = "/template/beetl/defaultExcelBuilderTemplate.html";

    private ExcelBuilder excelBuilder;
    /**
     * 标题
     */
    private List<String> titles;
    /**
     * sheetName
     */
    private String sheetName;
    /**
     * 字段展示顺序
     */
    private List<String> fieldDisplayOrder;

    private DefaultExcelBuilder() {
        this.excelBuilder = new BeetlExcelBuilder();
    }

    public static DefaultExcelBuilder getInstance() {
        return new DefaultExcelBuilder();
    }

    public DefaultExcelBuilder titles(List<String> titles) {
        this.titles = titles;
        return this;
    }

    public DefaultExcelBuilder sheetName(String sheetName) {
        this.sheetName = Objects.isNull(sheetName) ? "sheet" : sheetName;
        return this;
    }

    public DefaultExcelBuilder fieldDisplayOrder(List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    public Workbook build(List<?> data) {
        if (Objects.isNull(fieldDisplayOrder) || fieldDisplayOrder.isEmpty()) {
            throw new IllegalArgumentException("FieldDisplayOrder is necessary");
        }
        this.selfAdaption();
        // 设置标题
        Map<String, Object> renderData = new HashMap<>();
        renderData.put("titles", titles);

        renderData.put("sheetName", sheetName);

        if (Objects.isNull(data) || data.isEmpty()) {
            log.info("No valid data exists");
            return excelBuilder.template(DEFAULT_TEMPLATE_PATH).build(renderData);
        }
        Optional<?> findResult = data.stream().filter(Objects::nonNull).findFirst();
        if (!findResult.isPresent()) {
            log.info("No valid data exists");
            return excelBuilder.template(DEFAULT_TEMPLATE_PATH).build(renderData);
        }
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(findResult.get().getClass());
        List<Field> sortedFields = fieldDisplayOrder.stream()
                .map(classFieldContainer::getFieldByFieldName)
                .collect(Collectors.toList());
        if (sortedFields.isEmpty()) {
            log.info("The specified field mapping does not exist");
            return excelBuilder.template(DEFAULT_TEMPLATE_PATH).build(renderData);
        }
        List<List<Object>> contents = data.stream().map(d ->
                sortedFields.stream().map(field -> ReflectUtil.getFieldValue(d, field)).collect(Collectors.toList()))
                .collect(Collectors.toList());

        renderData.put("contents", contents);
        return excelBuilder.template(DEFAULT_TEMPLATE_PATH).build(renderData);
    }

    /**
     * 展示字段order与标题title长度一致性自适应
     */
    private void selfAdaption() {
        if (Objects.isNull(titles) || titles.isEmpty()) {
            return;
        }
        if (fieldDisplayOrder.size() < titles.size()) {
            for (int i = 0, size = titles.size() - fieldDisplayOrder.size(); i < size; i++) {
                fieldDisplayOrder.add(null);
            }
        } else {
            for (int i = 0, size = fieldDisplayOrder.size() - titles.size(); i < size; i++) {
                titles.add(null);
            }
        }
    }
}
