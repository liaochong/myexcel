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
import com.github.liaochong.html2excel.core.cache.Cache;
import com.github.liaochong.html2excel.core.cache.DefaultCache;
import com.github.liaochong.html2excel.core.parallel.ParallelContainer;
import com.github.liaochong.html2excel.core.parser.Table;
import com.github.liaochong.html2excel.core.parser.Td;
import com.github.liaochong.html2excel.core.parser.Tr;
import com.github.liaochong.html2excel.core.reflect.ClassFieldContainer;
import com.github.liaochong.html2excel.utils.ReflectUtil;
import com.github.liaochong.html2excel.utils.StringUtil;
import com.github.liaochong.html2excel.utils.TdUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * 默认excel创建者
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class DefaultExcelBuilder {

    private static final Cache<String, DateTimeFormatter> DATETIME_FORMATTER_CONTAINER = new DefaultCache<>();

    private static final Map<String, String> COMMON_STYLE = new HashMap<>();

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
    private WorkbookType workbookType;
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

    static {
        COMMON_STYLE.put("border-bottom-style", "thin");
        COMMON_STYLE.put("border-left-style", "thin");
        COMMON_STYLE.put("border-right-style", "thin");
    }

    private DefaultExcelBuilder() {
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

    /**
     * 设置workbookType为SXSSFWorkbook的内存数据保有量
     *
     * @param rowAccessWindowSize 内存数据保有量
     * @return HtmlToExcelFactory
     */
    public DefaultExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        this.rowAccessWindowSize = rowAccessWindowSize;
        return this;
    }

    /**
     * 设置workbook类型
     *
     * @param workbookType 工作簿类型
     * @return HtmlToExcelFactory
     */
    public DefaultExcelBuilder workbookType(WorkbookType workbookType) {
        this.workbookType = workbookType;
        return this;
    }

    public Workbook build(List<?> data) {
        if (Objects.isNull(data) || data.isEmpty()) {
            log.info("No valid data exists");
            return new HtmlToExcelFactory().build(Collections.emptyList());
        }
        Optional<?> findResult = data.stream().filter(Objects::nonNull).findFirst();
        if (!findResult.isPresent()) {
            log.info("No valid data exists");
            return new HtmlToExcelFactory().build(Collections.emptyList());
        }
        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(findResult.get().getClass());
        List<Field> sortedFields = getSortedFieldsAndSetting(classFieldContainer);

        if (sortedFields.isEmpty()) {
            log.info("The specified field mapping does not exist");
            return new HtmlToExcelFactory().build(Collections.emptyList());
        }
        List<List<Object>> contents = getRenderContent(data, sortedFields);

        List<Table> tableList = new ArrayList<>();
        tableList.add(this.createTable(contents));
        workbookType = Objects.nonNull(workbookType) ? workbookType : WorkbookType.XLSX;
        HtmlToExcelFactory htmlToExcelFactory = new HtmlToExcelFactory();
        htmlToExcelFactory.rowAccessWindowSize(rowAccessWindowSize).workbookType(workbookType);
        return htmlToExcelFactory.build(tableList);
    }

    public DefaultExcelBuilder startStreamBuild(Class<?> clazz) {
        this.startStreamBuild(clazz, HtmlToExcelStreamFactory.DEFAULT_WAIT_SIZE);
        return this;
    }

    public DefaultExcelBuilder startStreamBuild(Class<?> clazz, int waitQueueSize) {
        Objects.requireNonNull(clazz);
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(waitQueueSize);

        ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(clazz);
        sortedFields = getSortedFieldsAndSetting(classFieldContainer);

        Table table = new Table();
        table.setCaption(sheetName);

        htmlToExcelStreamFactory.start(table);

        Tr head = this.getThead();
        if (Objects.isNull(head)) {
            return this;
        }
        List<Tr> headList = new ArrayList<>();
        headList.add(head);
        htmlToExcelStreamFactory.append(headList);
        return this;
    }

    public void append(List<?> data) {
        if (Objects.isNull(data) || data.isEmpty()) {
            return;
        }
        List<List<Object>> contents = getRenderContent(data, sortedFields);
        List<Tr> trList = this.getTbody(0, contents);
        htmlToExcelStreamFactory.append(trList);
    }

    public Workbook streamBuild() {
        return htmlToExcelStreamFactory.build();
    }

    /**
     * 获取排序后字段并设置标题、workbookType等
     *
     * @param classFieldContainer classFieldContainer
     * @return Field
     */
    private List<Field> getSortedFieldsAndSetting(ClassFieldContainer classFieldContainer) {
        ExcelTable excelTable = classFieldContainer.getClazz().getAnnotation(ExcelTable.class);

        boolean excludeParent = false;
        boolean includeAllField = false;
        if (Objects.nonNull(excelTable)) {
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

        List<String> titles = new ArrayList<>();
        List<Field> sortedFields = preelectionFields.stream()
                .filter(field -> !field.isAnnotationPresent(ExcludeColumn.class))
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
                    if (Objects.isNull(excelColumn)) {
                        titles.add(null);
                    } else {
                        titles.add(excelColumn.title());
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
     * 获取并且转换字段值
     *
     * @param data  数据
     * @param field 对应字段
     * @return 结果
     */
    private Object getAndConvertFieldValue(Object data, Field field) {
        Object result = ReflectUtil.getFieldValue(data, field);
        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
        if (Objects.isNull(excelColumn) || Objects.isNull(result)) {
            return result;
        }
        // 时间格式化
        String dateFormatPattern = excelColumn.dateFormatPattern();
        if (StringUtil.isNotBlank(dateFormatPattern)) {
            Class<?> fieldType = field.getType();
            if (fieldType == LocalDateTime.class) {
                LocalDateTime localDateTime = (LocalDateTime) result;
                DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
                return formatter.format(localDateTime);
            } else if (fieldType == LocalDate.class) {
                LocalDate localDate = (LocalDate) result;
                DateTimeFormatter formatter = getDateTimeFormatter(dateFormatPattern);
                return formatter.format(localDate);
            } else if (fieldType == Date.class) {
                Date date = (Date) result;
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormatPattern);
                return simpleDateFormat.format(date);
            }
        }
        return result;
    }

    /**
     * 获取时间格式化
     *
     * @param dateFormat 时间格式化
     * @return DateTimeFormatter
     */
    private DateTimeFormatter getDateTimeFormatter(String dateFormat) {
        DateTimeFormatter formatter = DATETIME_FORMATTER_CONTAINER.get(dateFormat);
        if (Objects.isNull(formatter)) {
            formatter = DateTimeFormatter.ofPattern(dateFormat);
            DATETIME_FORMATTER_CONTAINER.cache(dateFormat, formatter);
        }
        return formatter;
    }

    /**
     * 获取需要被渲染的内容
     *
     * @param data         数据集合
     * @param sortedFields 排序字段
     * @return 结果集
     */
    private List<List<Object>> getRenderContent(List<?> data, List<Field> sortedFields) {
        List<ParallelContainer> resolvedDataContainers = IntStream.range(0, data.size()).parallel().mapToObj(index -> {
            List<Object> resolvedDataList = sortedFields.stream()
                    .map(field -> this.getAndConvertFieldValue(data.get(index), field))
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
     * 获取table
     *
     * @param contents 渲染内容
     * @return table
     */
    private Table createTable(List<List<Object>> contents) {
        Table table = new Table();
        table.setCaption(sheetName);

        table.setTrList(new ArrayList<>());

        Tr thead = this.getThead();
        boolean hasTitles = Objects.nonNull(thead);
        if (hasTitles) {
            table.getTrList().add(thead);
        }
        List<Tr> contentTrList = this.getTbody(hasTitles ? 1 : 0, contents);
        table.getTrList().addAll(contentTrList);
        return table;
    }

    /**
     * 获取thead
     *
     * @return tr
     */
    private Tr getThead() {
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
            tr.getColWidthMap().put(index, TdUtil.getStringWidth(td.getContent()));
            return td;
        }).collect(Collectors.toList());
        tr.setTdList(ths);
        return tr;
    }


    private List<Tr> getTbody(int shift, List<List<Object>> contents) {
        Map<String, String> oddTdStyle = new HashMap<>(COMMON_STYLE);
        oddTdStyle.put("background-color", "#f6f8fa");
        // 偏移量
        return IntStream.range(0, contents.size()).parallel().mapToObj(index -> {
            int trIndex = index + shift;
            Tr tr = new Tr(trIndex);
            List<Object> dataList = contents.get(index);
            Map<Integer, Integer> colMaxWidthMap = new HashMap<>(dataList.size());
            tr.setColWidthMap(colMaxWidthMap);
            Map<String, String> tdStyle = tr.getIndex() % 2 == 0 ? COMMON_STYLE : oddTdStyle;
            List<Td> tdList = IntStream.range(0, dataList.size()).mapToObj(i -> {
                Td td = new Td();
                td.setRow(trIndex);
                td.setRowBound(trIndex);
                td.setCol(i);
                td.setColBound(i);
                td.setContent(Objects.isNull(dataList.get(i)) ? null : String.valueOf(dataList.get(i)));
                td.setStyle(tdStyle);
                tr.getColWidthMap().put(i, TdUtil.getStringWidth(td.getContent()));
                return td;
            }).collect(Collectors.toList());
            tr.setTdList(tdList);
            contents.set(index, null);
            return tr;
        }).collect(Collectors.toList());
    }

}
