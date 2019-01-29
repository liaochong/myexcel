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

import com.github.liaochong.html2excel.core.annotation.ExcelColumn;
import com.github.liaochong.html2excel.core.annotation.ExcelTable;
import com.github.liaochong.html2excel.core.annotation.ExcludeColumn;
import com.github.liaochong.html2excel.core.converter.ConverterContext;
import com.github.liaochong.html2excel.core.converter.DateTimeConverter;
import com.github.liaochong.html2excel.core.parallel.ParallelContainer;
import com.github.liaochong.html2excel.core.parser.Table;
import com.github.liaochong.html2excel.core.parser.Td;
import com.github.liaochong.html2excel.core.parser.Tr;
import com.github.liaochong.html2excel.core.reflect.ClassFieldContainer;
import com.github.liaochong.html2excel.core.style.FontStyle;
import com.github.liaochong.html2excel.utils.ReflectUtil;
import com.github.liaochong.html2excel.utils.StringUtil;
import com.github.liaochong.html2excel.utils.TdUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.ExecutorService;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelBuilder implements SimpleExcelBuilder, SimpleStreamExcelBuilder {

    /**
     * 一般单元格样式
     */
    private Map<String, String> commonTdStyle;
    /**
     * 奇数行单元格样式
     */
    private Map<String, String> oddTdStyle;
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
    private WorkbookType workbookType = WorkbookType.XLSX;
    /**
     * 内存数据保有量
     */
    private int rowAccessWindowSize;
    /**
     * 流工厂
     */
    private HtmlToExcelStreamFactory htmlToExcelStreamFactory;
    /**
     * 已排序字段
     */
    private List<Field> sortedFields;
    /**
     * 线程池
     */
    private ExecutorService executorService;
    /**
     * 设置需要渲染的数据的类类型
     */
    private Class<?> dataType;

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
    public static DefaultExcelBuilder of(Class<?> dataType) {
        Objects.requireNonNull(dataType);
        DefaultExcelBuilder defaultExcelBuilder = new DefaultExcelBuilder();
        defaultExcelBuilder.dataType = dataType;
        return defaultExcelBuilder;
    }

    @Override
    public DefaultExcelBuilder titles(List<String> titles) {
        Objects.requireNonNull(titles);
        this.titles = titles;
        return this;
    }

    @Override
    public DefaultExcelBuilder sheetName(String sheetName) {
        Objects.requireNonNull(sheetName);
        this.sheetName = sheetName;
        return this;
    }

    @Override
    public DefaultExcelBuilder fieldDisplayOrder(List<String> fieldDisplayOrder) {
        Objects.requireNonNull(fieldDisplayOrder);
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    @Override
    public DefaultExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        if (rowAccessWindowSize <= 0) {
            throw new IllegalArgumentException("RowAccessWindowSize must be greater than 0");
        }
        this.rowAccessWindowSize = rowAccessWindowSize;
        return this;
    }

    @Override
    public DefaultExcelBuilder workbookType(WorkbookType workbookType) {
        Objects.requireNonNull(workbookType);
        this.workbookType = workbookType;
        return this;
    }

    @Override
    public DefaultExcelBuilder threadPool(ExecutorService executorService) {
        Objects.requireNonNull(executorService);
        this.executorService = executorService;
        return this;
    }

    @Override
    public Workbook build(List<?> data, Class<?>... groups) {
        HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();
        List<Table> tableList = new ArrayList<>();
        if (Objects.isNull(dataType)) {
            if (Objects.isNull(data) || data.isEmpty()) {
                log.info("No valid data exists");
                return htmlToExcelFactory.build(this.getTableWithHeader());
            }
            Optional<?> findResult = data.stream().filter(Objects::nonNull).findFirst();
            if (!findResult.isPresent()) {
                log.info("No valid data exists");
                return htmlToExcelFactory.build(this.getTableWithHeader());
            }
            ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(findResult.get().getClass());
            List<Field> sortedFields = getFilteredFields(classFieldContainer, groups);

            if (sortedFields.isEmpty()) {
                log.info("The specified field mapping does not exist");
                return htmlToExcelFactory.build(this.getTableWithHeader());
            }
            List<List<Object>> contents = getRenderContent(data, sortedFields);

            this.initStyleMap();

            Table table = this.createTable();
            Tr thead = this.createThead();
            if (Objects.nonNull(thead)) {
                table.getTrList().add(thead);
            }
            List<Tr> tbody = this.createTbody(contents, Objects.isNull(thead) ? 0 : 1);
            table.getTrList().addAll(tbody);
            tableList.add(table);
        } else {
            ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
            List<Field> sortedFields = getFilteredFields(classFieldContainer, groups);

            if (sortedFields.isEmpty()) {
                log.info("The specified field mapping does not exist");
                return htmlToExcelFactory.build(Collections.emptyList());
            }

            Table table = this.createTable();
            Tr thead = this.createThead();
            if (Objects.nonNull(thead)) {
                table.getTrList().add(thead);
            }
            tableList.add(table);

            if (Objects.isNull(data) || data.isEmpty()) {
                log.info("No valid data exists");
                return htmlToExcelFactory.build(tableList);
            }

            this.initStyleMap();

            List<List<Object>> contents = getRenderContent(data, sortedFields);
            List<Tr> tbody = this.createTbody(contents, Objects.isNull(thead) ? 0 : 1);
            table.getTrList().addAll(tbody);
        }
        htmlToExcelFactory.rowAccessWindowSize(rowAccessWindowSize).workbookType(workbookType);
        return htmlToExcelFactory.build(tableList);
    }

    /**
     * 获取只有head的table
     *
     * @return table集合
     */
    private List<Table> getTableWithHeader() {
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
     * 流式构建启动，包含一些初始化操作，等待队列容量采用CPU核心数目
     *
     * @param groups 分组
     * @return DefaultExcelBuilder
     */
    public DefaultExcelBuilder start(Class<?>... groups) {
        this.start(HtmlToExcelStreamFactory.DEFAULT_WAIT_SIZE, groups);
        return this;
    }

    @Override
    public DefaultExcelBuilder start(int waitQueueSize, Class<?>... groups) {
        Objects.requireNonNull(dataType);
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(waitQueueSize, executorService);

        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
        sortedFields = getFilteredFields(classFieldContainer, groups);

        this.initStyleMap();
        Table table = this.createTable();
        htmlToExcelStreamFactory.start(table);

        Tr head = this.createThead();
        if (Objects.isNull(head)) {
            return this;
        }
        List<Tr> headList = new ArrayList<>();
        headList.add(head);
        htmlToExcelStreamFactory.append(headList);
        return this;
    }

    @Override
    public void append(List<?> data) {
        if (Objects.isNull(data) || data.isEmpty()) {
            return;
        }
        List<List<Object>> contents = getRenderContent(data, sortedFields);
        List<Tr> trList = this.createTbody(contents, 0);
        htmlToExcelStreamFactory.append(trList);
    }

    @Override
    public Workbook build() {
        return htmlToExcelStreamFactory.build();
    }

    /**
     * 获取排序后字段并设置标题、workbookType等
     *
     * @param classFieldContainer classFieldContainer
     * @param groups              分组
     * @return Field
     */
    private List<Field> getFilteredFields(ClassFieldContainer classFieldContainer, Class<?>... groups) {
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
    private List<List<Object>> getRenderContent(List<?> data, List<Field> sortedFields) {
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

    /**
     * 创建table
     *
     * @return table
     */
    private Table createTable() {
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
    private Tr createThead() {
        boolean hasTitles = Objects.nonNull(titles) && !titles.isEmpty();
        if (!hasTitles) {
            return null;
        }
        Map<String, String> thStyle = new HashMap<>();
        thStyle.put("font-weight", "bold");
        thStyle.put("font-size", "14");
        thStyle.put("text-align", "center");
        thStyle.put("vertical-align", "center");
        thStyle.put("border-bottom-style", "thin");
        thStyle.put("border-left-style", "thin");
        thStyle.put("border-right-style", "thin");

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
    private List<Tr> createTbody(List<List<Object>> contents, int shift) {
        return IntStream.range(0, contents.size()).parallel().mapToObj(index -> {
            int trIndex = index + shift;
            Tr tr = new Tr(trIndex);
            List<Object> dataList = contents.get(index);
            Map<Integer, Integer> colMaxWidthMap = new HashMap<>(dataList.size());
            tr.setColWidthMap(colMaxWidthMap);
            Map<String, String> tdStyle = tr.getIndex() % 2 == 0 ? commonTdStyle : oddTdStyle;
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
    private void initStyleMap() {
        commonTdStyle = new HashMap<>(3);
        commonTdStyle.put("border-bottom-style", "thin");
        commonTdStyle.put("border-left-style", "thin");
        commonTdStyle.put("border-right-style", "thin");

        oddTdStyle = new HashMap<>(4);
        oddTdStyle.put("background-color", "#f6f8fa");
        oddTdStyle.putAll(commonTdStyle);
    }
}
