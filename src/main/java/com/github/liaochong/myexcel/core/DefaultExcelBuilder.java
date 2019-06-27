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

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
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

    private DefaultExcelBuilder() {
    }

    /**
     * 获取实例，已废弃，请使用of方法代替
     *
     * @return DefaultExcelBuilder
     */
    @Deprecated
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

    @Override
    public Workbook build(List<?> data, Class<?>... groups) {
        HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();
        htmlToExcelFactory.rowAccessWindowSize(rowAccessWindowSize).workbookType(workbookType).autoWidthStrategy(autoWidthStrategy);
        List<Table> tableList = new ArrayList<>();
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
                return htmlToExcelFactory.build(this.getTableWithHeader());
            }
            List<List<Pair<Class, Object>>> contents = getRenderContent(data, sortedFields);

            this.initStyleMap();

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

            this.initStyleMap();

            List<List<Pair<Class, Object>>> contents = getRenderContent(data, sortedFields);
            List<Tr> tbody = this.createTbody(contents, Objects.isNull(thead) ? 0 : thead.size());
            if (Objects.nonNull(thead)) {
                tbody.addAll(0, thead);
            }
            table.setTrList(tbody);
        }

        if (fixedTitles && titleLevel > 0) {
            FreezePane freezePane = new FreezePane(titleLevel, titles.size());
            htmlToExcelFactory.freezePanes(freezePane);
        }
        return htmlToExcelFactory.build(tableList, workbook);
    }
}
