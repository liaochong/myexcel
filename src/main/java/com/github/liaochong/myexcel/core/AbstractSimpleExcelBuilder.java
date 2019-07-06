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
import com.github.liaochong.myexcel.core.container.Pair;
import com.github.liaochong.myexcel.core.container.ParallelContainer;
import com.github.liaochong.myexcel.core.converter.WriteConverterContext;
import com.github.liaochong.myexcel.core.parser.ContentTypeEnum;
import com.github.liaochong.myexcel.core.parser.Table;
import com.github.liaochong.myexcel.core.parser.Td;
import com.github.liaochong.myexcel.core.parser.Tr;
import com.github.liaochong.myexcel.core.reflect.ClassFieldContainer;
import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import com.github.liaochong.myexcel.core.style.BackgroundStyle;
import com.github.liaochong.myexcel.core.style.BorderStyle;
import com.github.liaochong.myexcel.core.style.FontStyle;
import com.github.liaochong.myexcel.core.style.TextAlignStyle;
import com.github.liaochong.myexcel.core.style.WordBreakStyle;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import lombok.NonNull;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedList;
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
    protected List<String> titles;
    /**
     * sheetName
     */
    private String sheetName;
    /**
     * 字段展示顺序
     */
    protected List<String> fieldDisplayOrder;
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
     * 无样式
     */
    protected boolean noStyle;
    /**
     * 是否固定标题
     */
    protected boolean fixedTitles;
    /**
     * 自动宽度策略
     */
    protected AutoWidthStrategy autoWidthStrategy = AutoWidthStrategy.COMPUTE_AUTO_WIDTH;
    /**
     * 全局默认值
     */
    private String globalDefaultValue;
    /**
     * 默认值集合
     */
    private Map<Field, String> defaultValueMap;
    /**
     * 自定义宽度
     */
    private Map<Integer, Integer> customWidthMap;
    /**
     * 是否自动换行
     */
    private boolean wrapText = true;
    /**
     * 标题层级
     */
    protected int titleLevel = 0;
    /**
     * 标题分离器
     */
    private String titleSeparator = "->";


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
    public AbstractSimpleExcelBuilder workbookType(@NonNull WorkbookType workbookType) {
        this.workbookType = workbookType;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder noStyle() {
        this.noStyle = true;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        this.autoWidthStrategy = autoWidthStrategy;
        return this;
    }

    @Override
    public AbstractSimpleExcelBuilder fixedTitles() {
        this.fixedTitles = true;
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
        List<Tr> thead = this.createThead();
        if (Objects.nonNull(thead)) {
            table.getTrList().addAll(thead);
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
        table.setTrList(new LinkedList<>());
        return table;
    }

    /**
     * 创建标题行
     *
     * @return 标题行
     */
    protected List<Tr> createThead() {
        if (Objects.isNull(titles) || titles.isEmpty()) {
            return Collections.emptyList();
        }
        List<List<Td>> tdLists = new ArrayList<>();
        // 初始化位置信息
        for (int i = 0; i < titles.size(); i++) {
            String title = titles.get(i);
            if (Objects.isNull(title)) {
                continue;
            }
            List<Td> tds = new ArrayList<>();
            String[] multiTitles = title.split(titleSeparator);
            if (multiTitles.length > titleLevel) {
                titleLevel = multiTitles.length;
            }
            for (int j = 0; j < multiTitles.length; j++) {
                Td td = new Td();
                td.setTh(true);
                td.setCol(i);
                td.setRow(j);
                td.setContent(multiTitles[j]);
                tds.add(td);
            }
            tdLists.add(tds);
        }

        // 调整rowSpan
        for (List<Td> tdList : tdLists) {
            Td last = tdList.get(tdList.size() - 1);
            last.setRowSpan(titleLevel - last.getRow());
        }

        // 调整colSpan
        for (int i = 1; i < titleLevel; i++) {
            int level = i;
            Map<String, List<List<Td>>> groups = tdLists.stream()
                    .filter(list -> list.size() > level)
                    .collect(Collectors.groupingBy(list -> list.get(level - 1).getContent()));

            groups.forEach((k, v) -> {
                if (v.size() == 1) {
                    return;
                }
                List<Td> tds = groups.values().stream().flatMap(List::stream)
                        .map(list -> list.get(level - 1))
                        .sorted(Comparator.comparing(Td::getCol))
                        .collect(Collectors.toList());
                Td t = tds.get(0);
                t.setColSpan(v.size());
                for (int j = 1; j < tds.size(); j++) {
                    tds.get(j).setRow(-1);
                }
            });

        }

        Map<String, String> thStyle;
        if (noStyle) {
            thStyle = Collections.emptyMap();
        } else {
            thStyle = new HashMap<>(7);
            thStyle.put(FontStyle.FONT_WEIGHT, FontStyle.BOLD);
            thStyle.put(FontStyle.FONT_SIZE, "14");
            thStyle.put(TextAlignStyle.TEXT_ALIGN, TextAlignStyle.CENTER);
            thStyle.put(TextAlignStyle.VERTICAL_ALIGN, TextAlignStyle.MIDDLE);
            thStyle.put(BorderStyle.BORDER_BOTTOM_STYLE, BorderStyle.THIN);
            thStyle.put(BorderStyle.BORDER_LEFT_STYLE, BorderStyle.THIN);
            thStyle.put(BorderStyle.BORDER_RIGHT_STYLE, BorderStyle.THIN);
        }

        Map<Integer, List<Td>> rowTds = tdLists.stream().flatMap(List::stream).filter(td -> td.getRow() > -1).collect(Collectors.groupingBy(Td::getRow));
        List<Tr> trs = new ArrayList<>();
        boolean isComputeAutoWidth = AutoWidthStrategy.isComputeAutoWidth(autoWidthStrategy);
        rowTds.forEach((k, v) -> {
            Tr tr = new Tr(k);
            tr.setColWidthMap(isComputeAutoWidth ? new HashMap<>(titles.size()) : Collections.emptyMap());
            List<Td> tds = v.stream().sorted(Comparator.comparing(Td::getCol))
                    .peek(td -> {
                        td.setStyle(thStyle);
                        if (isComputeAutoWidth) {
                            tr.getColWidthMap().put(td.getCol(), TdUtil.getStringWidth(td.getContent(), 0.25));
                        }
                    })
                    .collect(Collectors.toList());
            tr.setTdList(tds);
            trs.add(tr);
        });
        return trs;
    }

    /**
     * 创建内容行
     *
     * @param contents 内容集合
     * @param shift    行序号偏移量
     * @return 内容行集合
     */
    protected List<Tr> createTbody(List<List<Pair<? extends Class, ?>>> contents, int shift) {
        List<Tr> result = IntStream.range(0, contents.size()).parallel().mapToObj(index ->
                this.createTr(contents.get(index), index, shift)
        ).collect(Collectors.toCollection(LinkedList::new));
        contents.clear();
        return result;
    }

    /**
     * 创建内容行
     *
     * @param contents 内容集合
     * @param index    索引
     * @param shift    行序号偏移量
     * @return 内容行
     */
    protected Tr createTr(List<Pair<? extends Class, ?>> contents, int index, int shift) {
        boolean isComputeAutoWidth = AutoWidthStrategy.isComputeAutoWidth(autoWidthStrategy);
        boolean isCustomWidth = AutoWidthStrategy.isCustomWidth(autoWidthStrategy);
        int trIndex = index + shift;
        Tr tr = new Tr(trIndex);
        tr.setColWidthMap((isComputeAutoWidth || isCustomWidth) ? new HashMap<>(contents.size()) : Collections.emptyMap());
        Map<String, String> tdStyle = (index & 1) == 0 ? commonTdStyle : evenTdStyle;
        List<Td> tdList = IntStream.range(0, contents.size()).mapToObj(i -> {
            Td td = new Td();
            td.setRow(trIndex);
            td.setCol(i);

            Pair<? extends Class, ?> pair = contents.get(i);
            td.setContent(Objects.isNull(pair.getValue()) ? null : String.valueOf(pair.getValue()));
            Class fieldType = pair.getKey();
            setTdContentType(td, fieldType);
            td.setStyle(tdStyle);
            if (isComputeAutoWidth) {
                tr.getColWidthMap().put(i, TdUtil.getStringWidth(td.getContent()));
            }
            return td;
        }).collect(Collectors.toList());
        if (isCustomWidth) {
            customWidthMap.forEach((k, v) -> tr.getColWidthMap().put(k, v));
        }
        tr.setTdList(tdList);
        return tr;

    }

    private void setTdContentType(Td td, Class fieldType) {
        if (String.class == fieldType) {
            return;
        }
        if (Boolean.class == fieldType || boolean.class == fieldType) {
            td.setTdContentType(ContentTypeEnum.BOOLEAN);
            return;
        }
        if (fieldType == Double.class || fieldType == double.class
                || fieldType == Float.class || fieldType == float.class
                || fieldType == Long.class || fieldType == long.class
                || fieldType == Integer.class || fieldType == int.class
                || fieldType == Short.class || fieldType == short.class
                || fieldType == Byte.class || fieldType == byte.class
                || fieldType == BigDecimal.class) {
            td.setTdContentType(ContentTypeEnum.DOUBLE);
        }
    }

    /**
     * 初始化单元格样式
     */
    protected void initStyleMap() {
        if (noStyle) {
            commonTdStyle = evenTdStyle = Collections.emptyMap();
        } else {
            commonTdStyle = new HashMap<>(3);
            commonTdStyle.put(BorderStyle.BORDER_BOTTOM_STYLE, BorderStyle.THIN);
            commonTdStyle.put(BorderStyle.BORDER_LEFT_STYLE, BorderStyle.THIN);
            commonTdStyle.put(BorderStyle.BORDER_RIGHT_STYLE, BorderStyle.THIN);
            commonTdStyle.put(TextAlignStyle.VERTICAL_ALIGN, TextAlignStyle.MIDDLE);
            if (wrapText) {
                commonTdStyle.put(WordBreakStyle.WORD_BREAK, WordBreakStyle.BREAK_ALL);
            }

            evenTdStyle = new HashMap<>(4);
            evenTdStyle.put(BackgroundStyle.BACKGROUND_COLOR, "#f6f8fa");
            evenTdStyle.putAll(commonTdStyle);
        }
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
        boolean ignoreStaticFields = true;
        if (excelTableExist) {
            setWorkbookWithExcelTableAnnotation(excelTable);
            excludeParent = excelTable.excludeParent();
            includeAllField = excelTable.includeAllField();
            if (!excelTable.defaultValue().isEmpty()) {
                globalDefaultValue = excelTable.defaultValue();
            }
            wrapText = excelTable.wrapText();
            titleSeparator = excelTable.titleSeparator();
            ignoreStaticFields = excelTable.ignoreStaticFields();
        }
        List<Field> preElectionFields = this.getPreElectionFields(classFieldContainer, excludeParent, includeAllField);
        if (ignoreStaticFields) {
            preElectionFields = preElectionFields.stream()
                    .filter(field -> !Modifier.isStatic(field.getModifiers()))
                    .collect(Collectors.toList());
        }
        List<Class<?>> selectedGroupList = Objects.nonNull(groups) ? Arrays.stream(groups).filter(Objects::nonNull).collect(Collectors.toList()) : Collections.emptyList();
        boolean useFieldNameAsTitle = excelTableExist && excelTable.useFieldNameAsTitle();
        List<String> titles = new ArrayList<>(preElectionFields.size());
        List<Field> sortedFields = preElectionFields.stream()
                .filter(field -> !field.isAnnotationPresent(ExcludeColumn.class) && filterFields(selectedGroupList, field))
                .sorted(this::sortFields)
                .collect(Collectors.toList());
        defaultValueMap = new HashMap<>(preElectionFields.size());
        customWidthMap = new HashMap<>(sortedFields.size());

        boolean needToAddTitle = Objects.isNull(this.titles);
        for (int i = 0, size = sortedFields.size(); i < size; i++) {
            Field field = sortedFields.get(i);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (Objects.nonNull(excelColumn)) {
                if (needToAddTitle) {
                    if (useFieldNameAsTitle && excelColumn.title().isEmpty()) {
                        titles.add(field.getName());
                    } else {
                        titles.add(excelColumn.title());
                    }
                }
                if (!excelColumn.defaultValue().isEmpty()) {
                    defaultValueMap.put(field, excelColumn.defaultValue());
                }
                if (excelColumn.width() > 0) {
                    customWidthMap.put(i, excelColumn.width());
                }
            } else {
                if (needToAddTitle) {
                    if (useFieldNameAsTitle) {
                        titles.add(field.getName());
                    } else {
                        titles.add(null);
                    }
                }
            }
        }

        boolean hasTitle = titles.stream().anyMatch(StringUtil::isNotBlank);
        if (hasTitle) {
            this.titles = titles;
        }
        return sortedFields;
    }

    private List<Field> getPreElectionFields(ClassFieldContainer classFieldContainer, boolean excludeParent, boolean includeAllField) {
        if (Objects.nonNull(fieldDisplayOrder) && !fieldDisplayOrder.isEmpty()) {
            this.selfAdaption();
            return fieldDisplayOrder.stream()
                    .map(classFieldContainer::getFieldByName)
                    .collect(Collectors.toList());
        }
        List<Field> preElectionFields;
        if (includeAllField) {
            if (excludeParent) {
                preElectionFields = classFieldContainer.getDeclaredFields();
            } else {
                preElectionFields = classFieldContainer.getFields();
            }
            return preElectionFields;
        }
        if (excludeParent) {
            preElectionFields = classFieldContainer.getDeclaredFields().stream()
                    .filter(field -> field.isAnnotationPresent(ExcelColumn.class)).collect(Collectors.toList());
        } else {
            preElectionFields = classFieldContainer.getFieldsByAnnotation(ExcelColumn.class);
        }
        return preElectionFields;
    }

    private boolean filterFields(List<Class<?>> selectedGroupList, Field field) {
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
    }

    private int sortFields(Field field1, Field field2) {
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
        if (fieldDisplayOrder.size() > titles.size()) {
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
    protected List<List<Pair<? extends Class, ?>>> getRenderContent(List<?> data, List<Field> sortedFields) {
        List<ParallelContainer> resolvedDataContainers = IntStream.range(0, data.size()).parallel().mapToObj(index -> {
            List<Pair<? extends Class, ?>> resolvedDataList = this.getRenderContent(data.get(index), sortedFields);
            return new ParallelContainer<>(index, resolvedDataList);
        }).collect(Collectors.toCollection(LinkedList::new));
        data.clear();

        // 重排序
        return resolvedDataContainers.stream()
                .sorted(Comparator.comparing(ParallelContainer::getIndex))
                .map(ParallelContainer<List<Pair<? extends Class, ?>>>::getData).collect(Collectors.toCollection(LinkedList::new));
    }

    /**
     * 获取需要被渲染的内容
     *
     * @param data         数据集合
     * @param sortedFields 排序字段
     * @return 结果集
     */
    protected <T> List<Pair<? extends Class, ?>> getRenderContent(T data, List<Field> sortedFields) {
        return sortedFields.stream()
                .map(field -> {
                    Pair<? extends Class, Object> value = WriteConverterContext.convert(field, data);
                    if (Objects.nonNull(value.getValue())) {
                        return value;
                    }
                    String defaultValue = defaultValueMap.get(field);
                    if (Objects.nonNull(defaultValue)) {
                        return Pair.of(field.getType(), defaultValue);
                    }
                    if (Objects.nonNull(globalDefaultValue)) {
                        return Pair.of(field.getType(), globalDefaultValue);
                    }
                    return value;
                })
                .collect(Collectors.toCollection(LinkedList::new));
    }
}
