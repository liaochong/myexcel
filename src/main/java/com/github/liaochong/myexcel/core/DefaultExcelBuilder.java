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
package com.github.liaochong.myexcel.core;

import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelBuilder extends AbstractSimpleExcelBuilder {

    private Workbook workbook;

    private HtmlToExcelFactory htmlToExcelFactory;

    private DefaultExcelBuilder() {
    }

    /**
     * 获取实例
     *
     * @return DefaultExcelBuilder
     */
    public static DefaultExcelBuilder getInstance() {
        return new DefaultExcelBuilder();
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @return DefaultExcelBuilder
     */
    public static DefaultExcelBuilder of(@NonNull Class<?> dataType) {
        DefaultExcelBuilder defaultExcelBuilder = new DefaultExcelBuilder();
        defaultExcelBuilder.dataType = dataType;
        return defaultExcelBuilder;
    }

    public static DefaultExcelBuilder of(@NonNull Class<?> dataType, @NonNull Workbook workbook) {
        DefaultExcelBuilder defaultExcelBuilder = new DefaultExcelBuilder();
        defaultExcelBuilder.dataType = dataType;
        defaultExcelBuilder.workbook = workbook;
        return defaultExcelBuilder;
    }

    @SuppressWarnings("unchecked")
    @Override
    public Workbook build(List<?> data, Class<?>... groups) {
        htmlToExcelFactory = new HtmlToExcelFactory();
        try {
            htmlToExcelFactory.rowAccessWindowSize(rowAccessWindowSize).workbookType(workbookType).autoWidthStrategy(autoWidthStrategy);
            List<Table> tableList = new ArrayList<>();
            this.initStyleMap();
            if (data != null) {
                boolean isMapBuild = data.stream().anyMatch(d -> d instanceof Map);
                if (isMapBuild) {
                    return this.mapBuild((List<Map<String, Object>>) data, htmlToExcelFactory);
                }
            }
            if (Objects.isNull(dataType)) {
                if (Objects.isNull(data) || data.isEmpty()) {
                    log.info("No valid data exists");
                    return htmlToExcelFactory.build(this.getTableWithHeader(), workbook);
                }
                Optional<?> findResult = data.stream().filter(Objects::nonNull).findFirst();
                if (!findResult.isPresent()) {
                    log.info("No valid data exists");
                    return htmlToExcelFactory.build(this.getTableWithHeader(), workbook);
                }
                ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(findResult.get().getClass());
                List<Field> sortedFields = getFilteredFields(classFieldContainer, groups);

                if (sortedFields.isEmpty()) {
                    log.info("The specified field mapping does not exist");
                    return htmlToExcelFactory.build(this.getTableWithHeader(), workbook);
                }
                List<List<Pair<? extends Class, ?>>> contents = getRenderContent(data, sortedFields);

                Table table = this.createTable();
                List<Tr> thead = this.createThead();
                List<Tr> tbody = this.createTbody(contents, Objects.isNull(thead) ? 0 : thead.size());
                if (Objects.nonNull(thead)) {
                    tbody.addAll(0, thead);
                }
                table.setTrList(tbody);
                tableList.add(table);
            } else {
                ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
                List<Field> sortedFields = getFilteredFields(classFieldContainer, groups);

                Table table = this.createTable();
                List<Tr> thead = this.createThead();
                tableList.add(table);

                if (sortedFields.isEmpty()) {
                    if (Objects.nonNull(thead)) {
                        table.getTrList().addAll(thead);
                    }
                    log.info("The specified field mapping does not exist");
                    return htmlToExcelFactory.build(tableList, workbook);
                }

                if (Objects.isNull(data) || data.isEmpty()) {
                    if (Objects.nonNull(thead)) {
                        table.getTrList().addAll(thead);
                    }
                    log.info("No valid data exists");
                    return htmlToExcelFactory.build(tableList, workbook);
                }

                List<List<Pair<? extends Class, ?>>> contents = getRenderContent(data, sortedFields);
                List<Tr> tbody = this.createTbody(contents, Objects.isNull(thead) ? 0 : thead.size());
                if (Objects.nonNull(thead)) {
                    tbody.addAll(0, thead);
                }
                table.setTrList(tbody);
            }

            if (fixedTitles && titleLevel > 0) {
                FreezePane freezePane = new FreezePane(titleLevel, 0);
                htmlToExcelFactory.freezePanes(freezePane);
            }
            return htmlToExcelFactory.build(tableList, workbook);
        } catch (Exception e) {
            htmlToExcelFactory.closeWorkbook();
            throw new RuntimeException(e);
        }
    }

    private Workbook mapBuild(List<Map<String, Object>> data, HtmlToExcelFactory htmlToExcelFactory) {
        if (null == fieldDisplayOrder) {
            throw new IllegalArgumentException();
        }
        if (data == null || data.isEmpty()) {
            log.info("No valid data exists");
            return htmlToExcelFactory.build(this.getTableWithHeader(), workbook);
        }
        List<Tr> thead = this.createThead();
        List<Tr> tbody = new LinkedList<>();
        for (int i = 0, size = data.size(); i < size; i++) {
            Map<String, Object> d = data.get(i);
            if (d == null) {
                continue;
            }
            List<Pair<? extends Class, ?>> contents = new ArrayList<>(d.size());
            for (String fieldName : fieldDisplayOrder) {
                Object val = d.get(fieldName);
                contents.add(Pair.of(Objects.isNull(val) ? String.class : val.getClass(), val));
            }
            Tr tr = this.createTr(contents, i, thead.size());
            tbody.add(tr);
        }
        tbody.addAll(0, thead);
        Table table = this.createTable();
        table.setTrList(tbody);

        if (fixedTitles && titleLevel > 0) {
            FreezePane freezePane = new FreezePane(titleLevel, 0);
            htmlToExcelFactory.freezePanes(freezePane);
        }
        List<Table> tableList = new ArrayList<>();
        tableList.add(table);
        return htmlToExcelFactory.build(tableList, workbook);
    }

    @Override
    public void close() throws IOException {
        if (htmlToExcelFactory != null) {
            htmlToExcelFactory.closeWorkbook();
        }
    }
}
