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

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import com.github.liaochong.myexcel.core.annotation.ExcludeColumn;
import com.github.liaochong.myexcel.core.converter.ConverterContext;
import com.github.liaochong.myexcel.core.converter.DateTimeConverter;
import com.github.liaochong.myexcel.core.parallel.ParallelContainer;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Td;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.core.strategy.CellStyleStrategy;
import com.github.liaochong.myexcel.core.style.BackgroundStyle;
import com.github.liaochong.myexcel.core.style.BorderStyle;
import com.github.liaochong.myexcel.core.style.FontStyle;
import com.github.liaochong.myexcel.core.style.TextAlignStyle;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import lombok.NonNull;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
public abstract class AbstractSimpleExcelBuilder implements SimpleExcelBuilder {
    /**
     * 一般单元格样式
     */
    private Map<String, String> commonTdStyle;
    /**
     * 偶数行单元格样式
     */
    private Map<String, String> evenTdStyle;
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
    /**
     * excel workbook
     */
    protected WorkbookType workbookType = WorkbookType.XLSX;
    /**
     * 内存数据保有量
     */
    protected int rowAccessWindowSize;
    /**
     * 已排序字段
     */
    protected List<Field> filteredFields;
    /**
     * 设置需要渲染的数据的类类型
     */
    protected Class<?> dataType;
    /**
     * 样式策略
     */
    protected CellStyleStrategy cellStyleStrategy = CellStyleStrategy.DEFAULT_STYLE;

    @Override
    public AbstractSimpleExcelBuilder titles(@NonNull List<String> titles) {
        this.titles = titles;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder sheetName(@NonNull String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder cellStyleStrategy(@NonNull CellStyleStrategy cellStyleStrategy) {
        this.cellStyleStrategy = cellStyleStrategy;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder fieldDisplayOrder(@NonNull List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        if (rowAccessWindowSize <= 0) {
            throw new IllegalArgumentException("RowAccessWindowSize must be greater than 0");
        }
        this.rowAccessWindowSize = rowAccessWindowSize;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder workbookType(WorkbookType workbookType) {
        Objects.requireNonNull(workbookType);
        this.workbookType = workbookType;
        return this;
    }

    /**
     * 获取只有head的table
     *
     * @return table集合
     */
    protected List<Table> getTableWithHeader() {
        List<Table> tableList = new ArrayList<>();
        Table table = this.createTable();
        tableList.add(table);
        Tr thead = this.createThead();
        if (Objects.nonNull(thead)) {
            table.getTrList().add(thead);
        }
        return tableList;
    }

    /**
     * 创建table
     *
     * @return table
     */
    protected Table createTable() {
        Table table = new Table();
        table.setCaption(sheetName);
        table.setTrList(new ArrayList<>());
        return table;
    }

    /**
     * 创建标题行
     *
     * @return 标题行
     */
    protected Tr createThead() {
        boolean hasTitles = Objects.nonNull(titles) && !titles.isEmpty();
        if (!hasTitles) {
            return null;
        }
        Map<String, String> thStyle;
        if (CellStyleStrategy.isDefaultStyle(cellStyleStrategy)) {
            thStyle = new HashMap<>(7);
            thStyle.put(FontStyle.FONT_WEIGHT, FontStyle.BOLD);
            thStyle.put(FontStyle.FONT_SIZE, "14");
            thStyle.put(TextAlignStyle.TEXT_ALIGN, TextAlignStyle.CENTER);
            thStyle.put(TextAlignStyle.VERTICAL_ALIGN, TextAlignStyle.MIDDLE);
            thStyle.put(BorderStyle.BORDER_BOTTOM_STYLE, BorderStyle.THIN);
            thStyle.put(BorderStyle.BORDER_LEFT_STYLE, BorderStyle.THIN);
            thStyle.put(BorderStyle.BORDER_RIGHT_STYLE, BorderStyle.THIN);
        } else {
            thStyle = Collections.emptyMap();
        }

        Tr tr = new Tr(0);
        Map<Integer, Integer> colMaxWidthMap = new HashMap<>(titles.size());
        tr.setColWidthMap(colMaxWidthMap);

        List<Td> ths = IntStream.range(0, titles.size()).mapToObj(index -> {
            Td td = new Td();
            td.setTh(true);
            td.setRow(0);
            td.setRowBound(0);
            td.setCol(index);
            td.setColBound(index);
            td.setContent(titles.get(index));
            td.setStyle(thStyle);
            tr.getColWidthMap().put(index, TdUtil.getStringWidth(td.getContent(), FontStyle.FONT_SIZE_SHIFT));
            return td;
        }).collect(Collectors.toList());
        tr.setTdList(ths);
        return tr;
    }

    /**
     * 创建内容行
     *
     * @param contents 内容集合
     * @param shift    行序号偏移量
     * @return 内容行集合
     */
    protected List<Tr> createTbody(List<List<Object>> contents, int shift) {
        return IntStream.range(0, contents.size()).parallel().mapToObj(index -> {
            int trIndex = index + shift;
            Tr tr = new Tr(trIndex);
            List<Object> dataList = contents.get(index);
            Map<Integer, Integer> colMaxWidthMap = new HashMap<>(dataList.size());
            tr.setColWidthMap(colMaxWidthMap);
            Map<String, String> tdStyle;
            if (CellStyleStrategy.isDefaultStyle(cellStyleStrategy)) {
                tdStyle = (index & 1) == 0 ? commonTdStyle : evenTdStyle;
            } else {
                tdStyle = Collections.emptyMap();
            }
            List<Td> tdList = IntStream.range(0, dataList.size()).mapToObj(i -> {
                Td td = new Td();
                td.setRow(trIndex);
                td.setRowBound(trIndex);
                td.setCol(i);
                td.setColBound(i);
                td.setContent(Objects.isNull(dataList.get(i)) ? null : String.valueOf(dataList.get(i)));
                td.setStyle(tdStyle);
                tr.getColWidthMap().put(i, TdUtil.getStringWidth(td.getContent(), 0));
                return td;
            }).collect(Collectors.toList());
            tr.setTdList(tdList);
            contents.set(index, null);
            return tr;
        }).collect(Collectors.toList());
    }

    /**
     * 初始化单元格样式
     */
    protected void initStyleMap() {
        if (CellStyleStrategy.isNoStyle(cellStyleStrategy)) {
            return;
        }
        commonTdStyle = new HashMap<>(3);
        commonTdStyle.put(BorderStyle.BORDER_BOTTOM_STYLE, BorderStyle.THIN);
        commonTdStyle.put(BorderStyle.BORDER_LEFT_STYLE, BorderStyle.THIN);
        commonTdStyle.put(BorderStyle.BORDER_RIGHT_STYLE, BorderStyle.THIN);
        commonTdStyle.put(TextAlignStyle.VERTICAL_ALIGN, TextAlignStyle.MIDDLE);

        evenTdStyle = new HashMap<>(4);
        evenTdStyle.put(BackgroundStyle.BACKGROUND_COLOR, "#f6f8fa");
        evenTdStyle.putAll(commonTdStyle);
    }

    /**
     * 获取排序后字段并设置标题、workbookType等
     *
     * @param classFieldContainer classFieldContainer
     * @param groups              分组
     * @return Field
     */
    protected List<Field> getFilteredFields(ClassFieldContainer classFieldContainer, Class<?>... groups) {
        ExcelTable excelTable = classFieldContainer.getClazz().getAnnotation(ExcelTable.class);

        boolean excelTableExist = Objects.nonNull(excelTable);
        boolean excludeParent = false;
        boolean includeAllField = false;
        if (excelTableExist) {
            setWorkbookWithExcelTableAnnotation(excelTable);
            excludeParent = excelTable.excludeParent();
            includeAllField = excelTable.includeAllField();
        }

        List<Field> preelectionFields;
        if (includeAllField) {
            if (excludeParent) {
                preelectionFields = classFieldContainer.getDeclaredFields();
            } else {
                preelectionFields = classFieldContainer.getFields();
            }
        } else {
            if (excludeParent) {
                preelectionFields = classFieldContainer.getDeclaredFields().stream()
                        .filter(field -> field.isAnnotationPresent(ExcelColumn.class)).collect(Collectors.toList());
            } else {
                preelectionFields = classFieldContainer.getFieldsByAnnotation(ExcelColumn.class);
            }
            if (preelectionFields.isEmpty()) {
                if (Objects.isNull(fieldDisplayOrder) || fieldDisplayOrder.isEmpty()) {
                    throw new IllegalArgumentException("FieldDisplayOrder is necessary");
                }
                this.selfAdaption();
                return fieldDisplayOrder.stream()
                        .map(classFieldContainer::getFieldByName)
                        .collect(Collectors.toList());
            }
        }

        List<Class<?>> selectedGroupList = Objects.nonNull(groups) ? Arrays.stream(groups).filter(Objects::nonNull).collect(Collectors.toList()) : Collections.emptyList();
        boolean useFieldNameAsTitle = excelTableExist && excelTable.useFieldNameAsTitle();
        List<String> titles = new ArrayList<>();
        List<Field> sortedFields = preelectionFields.stream()
                .filter(field -> !field.isAnnotationPresent(ExcludeColumn.class))
                .filter(field -> {
                    if (selectedGroupList.isEmpty()) {
                        return true;
                    }
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (Objects.isNull(excelColumn)) {
                        return false;
                    }
                    Class<?>[] groupArr = excelColumn.groups();
                    if (groupArr.length == 0) {
                        return false;
                    }
                    List<Class<?>> reservedGroupList = Arrays.stream(groupArr).collect(Collectors.toList());
                    return reservedGroupList.stream().anyMatch(selectedGroupList::contains);
                })
                .sorted((field1, field2) -> {
                    ExcelColumn excelColumn1 = field1.getAnnotation(ExcelColumn.class);
                    ExcelColumn excelColumn2 = field2.getAnnotation(ExcelColumn.class);
                    if (Objects.isNull(excelColumn1) && Objects.isNull(excelColumn2)) {
                        return 0;
                    }
                    int defaultOrder = 0;
                    int order1 = defaultOrder;
                    if (Objects.nonNull(excelColumn1)) {
                        order1 = excelColumn1.order();
                    }
                    int order2 = defaultOrder;
                    if (Objects.nonNull(excelColumn2)) {
                        order2 = excelColumn2.order();
                    }
                    if (order1 == order2) {
                        return 0;
                    }
                    return order1 > order2 ? 1 : -1;
                })
                .peek(field -> {
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (Objects.nonNull(excelColumn)) {
                        if (useFieldNameAsTitle && excelColumn.title().isEmpty()) {
                            titles.add(field.getName());
                        } else {
                            titles.add(excelColumn.title());
                        }
                    } else {
                        if (useFieldNameAsTitle) {
                            titles.add(field.getName());
                        } else {
                            titles.add(null);
                        }
                    }
                })
                .collect(Collectors.toList());

        boolean hasTitle = titles.stream().anyMatch(StringUtil::isNotBlank);
        if (hasTitle) {
            this.titles = titles;
        }
        return sortedFields;
    }

    /**
     * 设置workbook
     *
     * @param excelTable excelTable
     */
    private void setWorkbookWithExcelTableAnnotation(ExcelTable excelTable) {
        if (Objects.isNull(workbookType)) {
            this.workbookType = excelTable.workbookType();
        }
        if (this.rowAccessWindowSize <= 0) {
            int rowAccessWindowSize = excelTable.rowAccessWindowSize();
            if (rowAccessWindowSize > 0) {
                this.rowAccessWindowSize = rowAccessWindowSize;
            }
        }
        if (StringUtil.isBlank(this.sheetName)) {
            String sheetName = excelTable.sheetName();
            if (StringUtil.isNotBlank(sheetName)) {
                this.sheetName = sheetName;
            }
        }
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

    /**
     * 获取需要被渲染的内容
     *
     * @param data         数据集合
     * @param sortedFields 排序字段
     * @return 结果集
     */
    protected List<List<Object>> getRenderContent(List<?> data, List<Field> sortedFields) {
        ConverterContext converterContext = ConverterContext.newInstance().registering(new DateTimeConverter());

        List<ParallelContainer> resolvedDataContainers = IntStream.range(0, data.size()).parallel().mapToObj(index -> {
            List<Object> resolvedDataList = sortedFields.stream()
                    .map(field -> converterContext.convert(field, data.get(index)))
                    .collect(Collectors.toList());
            data.set(index, null);
            return new ParallelContainer<>(index, resolvedDataList);
        }).collect(Collectors.toList());

        // 重排序
        return resolvedDataContainers.stream()
                .sorted(Comparator.comparing(ParallelContainer::getIndex))
                .map(ParallelContainer<List<Object>>::getData).collect(Collectors.toList());
    }
}
