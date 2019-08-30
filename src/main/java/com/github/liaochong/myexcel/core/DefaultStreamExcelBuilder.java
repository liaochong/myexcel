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
import com.github.liaochong.myexcel.core.constant.BooleanDropDownList;
import com.github.liaochong.myexcel.core.constant.DropDownList;
import com.github.liaochong.myexcel.core.constant.LinkEmail;
import com.github.liaochong.myexcel.core.constant.LinkUrl;
import com.github.liaochong.myexcel.core.constant.NumberDropDownList;
import com.github.liaochong.myexcel.core.container.Pair;
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
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.StyleUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
public class DefaultStreamExcelBuilder implements SimpleStreamExcelBuilder {
    /**
     * 一般单元格样式
     */
    private Map<String, String> commonTdStyle;
    /**
     * 偶数行单元格样式
     */
    private Map<String, String> evenTdStyle;
    /**
     * 超链接公共样式
     */
    private Map<String, String> linkCommonStyle;
    /**
     * 超链接偶数行样式
     */
    private Map<String, String> linkEvenStyle;
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
     * 已排序字段
     */
    private List<Field> filteredFields;
    /**
     * 设置需要渲染的数据的类类型
     */
    private Class<?> dataType;
    /**
     * 无样式
     */
    private boolean noStyle;
    /**
     * 是否固定标题
     */
    private boolean fixedTitles;
    /**
     * 自动宽度策略
     */
    private AutoWidthStrategy autoWidthStrategy;
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
    private int titleLevel = 0;
    /**
     * 标题分离器
     */
    private String titleSeparator = "->";
    /**
     * 自定义样式
     */
    private Map<String, Map<String, String>> customStyle = new HashMap<>();

    private Map<Integer, Integer> widths;
    /**
     * 线程池
     */
    private ExecutorService executorService;
    /**
     * 流工厂
     */
    private HtmlToExcelStreamFactory htmlToExcelStreamFactory;
    /**
     * workbook
     */
    private Workbook workbook;
    /**
     * 文件分割,excel容量
     */
    private int capacity;
    /**
     * path消费者
     */
    private Consumer<Path> pathConsumer;
    /**
     * 是否为奇数行
     */
    private boolean isOddRow = true;

    private DefaultStreamExcelBuilder() {
        noStyle = true;
        autoWidthStrategy = AutoWidthStrategy.NO_AUTO;
        workbookType = WorkbookType.SXLSX;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @return DefaultExcelBuilder
     */
    public static DefaultStreamExcelBuilder of(@NonNull Class<?> dataType) {
        DefaultStreamExcelBuilder defaultStreamExcelBuilder = new DefaultStreamExcelBuilder();
        defaultStreamExcelBuilder.dataType = dataType;
        return defaultStreamExcelBuilder;
    }

    /**
     * 获取实例，设定需要渲染的数据的类类型
     *
     * @param dataType 数据的类类型
     * @param workbook workbook
     * @return DefaultExcelBuilder
     */
    public static DefaultStreamExcelBuilder of(@NonNull Class<?> dataType, @NonNull Workbook workbook) {
        DefaultStreamExcelBuilder defaultStreamExcelBuilder = new DefaultStreamExcelBuilder();
        defaultStreamExcelBuilder.dataType = dataType;
        defaultStreamExcelBuilder.workbook = workbook;
        return defaultStreamExcelBuilder;
    }

    public static DefaultStreamExcelBuilder getInstance() {
        return new DefaultStreamExcelBuilder();
    }

    public DefaultStreamExcelBuilder titles(@NonNull List<String> titles) {
        this.titles = titles;
        return this;
    }

    public DefaultStreamExcelBuilder sheetName(@NonNull String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    public DefaultStreamExcelBuilder fieldDisplayOrder(@NonNull List<String> fieldDisplayOrder) {
        this.fieldDisplayOrder = fieldDisplayOrder;
        return this;
    }

    /**
     * SXSSF模式下设置窗口保留行数，采用默认值是最佳实践，不建议修改
     *
     * @param rowAccessWindowSize 内存保留行数
     * @return DefaultStreamExcelBuilder
     */
    @Deprecated
    public DefaultStreamExcelBuilder rowAccessWindowSize(int rowAccessWindowSize) {
        // do nothing
        return this;
    }

    public DefaultStreamExcelBuilder workbookType(@NonNull WorkbookType workbookType) {
        this.workbookType = workbookType;
        return this;
    }

    /**
     * 设置为无样式
     *
     * @return DefaultStreamExcelBuilder
     */
    public DefaultStreamExcelBuilder noStyle() {
        this.noStyle = true;
        return this;
    }

    public DefaultStreamExcelBuilder autoWidthStrategy(@NonNull AutoWidthStrategy autoWidthStrategy) {
        this.autoWidthStrategy = autoWidthStrategy;
        return this;
    }

    public DefaultStreamExcelBuilder fixedTitles() {
        this.fixedTitles = true;
        return this;
    }

    public DefaultStreamExcelBuilder widths(int... widths) {
        if (widths.length == 0) {
            return this;
        }
        this.widths = new HashMap<>(widths.length);
        for (int i = 0, size = widths.length; i < size; i++) {
            this.widths.put(i, widths[i]);
        }
        customWidthMap = this.widths;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder threadPool(ExecutorService executorService) {
        this.executorService = executorService;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder hasStyle() {
        this.noStyle = false;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder capacity(int capacity) {
        if (capacity < 1) {
            throw new IllegalArgumentException("Capacity must be greater than 0");
        }
        this.capacity = capacity;
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder pathConsumer(Consumer<Path> pathConsumer) {
        this.pathConsumer = pathConsumer;
        return this;
    }

    /**
     * 流式构建启动，包含一些初始化操作，等待队列容量采用CPU核心数目
     *
     * @param groups 分组
     * @return DefaultExcelBuilder
     */
    public DefaultStreamExcelBuilder start(Class<?>... groups) {
        this.start(HtmlToExcelStreamFactory.DEFAULT_WAIT_SIZE, groups);
        return this;
    }

    @Override
    public DefaultStreamExcelBuilder start(int waitQueueSize, Class<?>... groups) {
        htmlToExcelStreamFactory = new HtmlToExcelStreamFactory(waitQueueSize, executorService, pathConsumer, capacity);
        htmlToExcelStreamFactory.rowAccessWindowSize(rowAccessWindowSize).workbookType(workbookType).autoWidthStrategy(autoWidthStrategy);

        if (dataType != null) {
            ClassFieldContainer classFieldContainer = ReflectUtil.getAllFieldsOfClass(dataType);
            filteredFields = getFilteredFields(classFieldContainer, groups);
        }

        this.initStyleMap();
        Table table = this.createTable();
        htmlToExcelStreamFactory.start(table, workbook);

        List<Tr> head = this.createThead();
        if (Objects.isNull(head)) {
            return this;
        }
        htmlToExcelStreamFactory.appendTitles(head);
        return this;
    }

    @SuppressWarnings("unchecked")
    @Override
    public void append(List<?> data) {
        if (data == null || data.isEmpty()) {
            return;
        }
        boolean isMapBuild = data.stream().anyMatch(d -> d instanceof Map);
        if (isMapBuild) {
            for (Object datum : data) {
                Map<String, Object> d = (Map<String, Object>) datum;
                List<Pair<? extends Class, ?>> contents = new ArrayList<>(d.size());
                for (String fieldName : fieldDisplayOrder) {
                    Object val = d.get(fieldName);
                    contents.add(Pair.of(val == null ? String.class : val.getClass(), val));
                }
                Tr tr = this.createTr(contents);
                htmlToExcelStreamFactory.append(tr);
            }
        } else {
            for (Object datum : data) {
                List<Pair<? extends Class, ?>> contents = getRenderContent(datum, filteredFields);
                Tr tr = this.createTr(contents);
                htmlToExcelStreamFactory.append(tr);
            }
        }
    }

    @SuppressWarnings("unchecked")
    @Override
    public <T> void append(T data) {
        if (data == null) {
            return;
        }
        List<Pair<? extends Class, ?>> contents;
        if (data instanceof Map) {
            Map<String, Object> d = (Map<String, Object>) data;
            contents = new ArrayList<>(d.size());
            for (String fieldName : fieldDisplayOrder) {
                Object val = d.get(fieldName);
                contents.add(Pair.of(val == null ? String.class : val.getClass(), val));
            }
        } else {
            contents = getRenderContent(data, filteredFields);
        }
        Tr tr = this.createTr(contents);
        htmlToExcelStreamFactory.append(tr);
    }

    @Override
    public Workbook build() {
        if (fixedTitles && titleLevel > 0) {
            FreezePane freezePane = new FreezePane(titleLevel, 0);
            htmlToExcelStreamFactory.freezePanes(freezePane);
        }
        return htmlToExcelStreamFactory.build();
    }

    @Override
    public List<Path> buildAsPaths() {
        return htmlToExcelStreamFactory.buildAsPaths();
    }

    @Override
    public Path buildAsZip(String fileName) {
        return htmlToExcelStreamFactory.buildAsZip(fileName);
    }

    @Override
    public void close() throws IOException {
        if (htmlToExcelStreamFactory != null) {
            htmlToExcelStreamFactory.closeWorkbook();
        }
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
        List<Tr> thead = this.createThead();
        if (thead != null) {
            table.getTrList().addAll(thead);
        }
        return tableList;
    }

    /**
     * 创建table
     *
     * @return table
     */
    private Table createTable() {
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
    private List<Tr> createThead() {
        if (titles == null || titles.isEmpty()) {
            return Collections.emptyList();
        }
        List<List<Td>> tdLists = new ArrayList<>();
        // 初始化位置信息
        for (int i = 0; i < titles.size(); i++) {
            String title = titles.get(i);
            if (title == null) {
                continue;
            }
            List<Td> tds = new ArrayList<>();
            String[] multiTitles = title.split(titleSeparator);
            if (multiTitles.length > titleLevel) {
                titleLevel = multiTitles.length;
            }
            for (int j = 0; j < multiTitles.length; j++) {
                Td td = new Td(j, i);
                td.setTh(true);
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
        for (int i = 0; i < titleLevel; i++) {
            int level = i;
            Map<String, List<List<Td>>> groups = tdLists.stream()
                    .filter(list -> list.size() > level)
                    .collect(Collectors.groupingBy(list -> list.get(level).getContent()));
            groups.forEach((k, v) -> {
                if (v.size() == 1) {
                    return;
                }
                List<Td> tds = v.stream().map(list -> list.get(level))
                        .sorted(Comparator.comparing(Td::getCol))
                        .collect(Collectors.toList());

                List<List<Td>> subTds = new LinkedList<>();
                // 不同跨行分别处理
                Map<Integer, List<Td>> partitions = tds.stream().collect(Collectors.groupingBy(Td::getRowSpan));
                partitions.forEach((col, subTdList) -> {
                    // 区分开不连续列
                    int splitIndex = 0;
                    for (int j = 0, size = subTdList.size() - 1; j < size; j++) {
                        Td current = subTdList.get(j);
                        Td next = subTdList.get(j + 1);
                        if (current.getCol() + 1 != next.getCol()) {
                            List<Td> sub = subTdList.subList(splitIndex, j + 1);
                            splitIndex = j + 1;
                            if (sub.size() <= 1) {
                                continue;
                            }
                            subTds.add(sub);
                        }
                    }
                    subTds.add(subTdList.subList(splitIndex, subTdList.size()));
                });

                subTds.forEach(val -> {
                    if (val.size() == 1) {
                        return;
                    }
                    Td t = val.get(0);
                    t.setColSpan(val.size());
                    for (int j = 1; j < val.size(); j++) {
                        val.get(j).setRow(-1);
                    }
                });
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
                        // 自定义样式存在时采用自定义样式
                        if (!noStyle && !customStyle.isEmpty()) {
                            Map<String, String> style = customStyle.getOrDefault("title&" + td.getCol(), Collections.emptyMap());
                            td.setStyle(style);
                        } else {
                            td.setStyle(thStyle);
                        }
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
     * @return 内容行
     */
    private Tr createTr(List<Pair<? extends Class, ?>> contents) {
        boolean isComputeAutoWidth = AutoWidthStrategy.isComputeAutoWidth(autoWidthStrategy);
        boolean isCustomWidth = AutoWidthStrategy.isCustomWidth(autoWidthStrategy);
        Tr tr = new Tr(0);
        tr.setColWidthMap((isComputeAutoWidth || isCustomWidth) ? new HashMap<>(contents.size()) : Collections.emptyMap());
        Map<String, String> tdStyle = isOddRow ? commonTdStyle : evenTdStyle;
        Map<String, String> linkStyle = isOddRow ? linkCommonStyle : linkEvenStyle;
        isOddRow = !isOddRow;
        List<Td> tdList = IntStream.range(0, contents.size()).mapToObj(i -> {
            Td td = new Td(0, i);

            Pair<? extends Class, ?> pair = contents.get(i);
            td.setContent(pair.getValue() == null ? null : String.valueOf(pair.getValue()));
            Class fieldType = pair.getKey();
            setTdContentType(td, fieldType);
            if (!noStyle && !customStyle.isEmpty()) {
                Map<String, String> style = customStyle.getOrDefault("cell&" + i, Collections.emptyMap());
                td.setStyle(style);
            } else {
                if (ContentTypeEnum.isLink(td.getTdContentType())) {
                    td.setStyle(linkStyle);
                } else {
                    td.setStyle(tdStyle);
                }
            }
            if (isComputeAutoWidth) {
                tr.getColWidthMap().put(i, TdUtil.getStringWidth(td.getContent()));
            }
            return td;
        }).collect(Collectors.toList());
        if (isCustomWidth) {
            tr.setColWidthMap(customWidthMap);
        }
        tr.setTdList(tdList);
        return tr;

    }

    private void setTdContentType(Td td, Class fieldType) {
        if (String.class == fieldType) {
            return;
        }
        if (ReflectUtil.isNumber(fieldType)) {
            td.setTdContentType(ContentTypeEnum.DOUBLE);
            return;
        }
        if (ReflectUtil.isBool(fieldType)) {
            td.setTdContentType(ContentTypeEnum.BOOLEAN);
            return;
        }
        if (fieldType == DropDownList.class) {
            td.setTdContentType(ContentTypeEnum.DROP_DOWN_LIST);
            return;
        }
        if (fieldType == NumberDropDownList.class) {
            td.setTdContentType(ContentTypeEnum.NUMBER_DROP_DOWN_LIST);
            return;
        }
        if (fieldType == BooleanDropDownList.class) {
            td.setTdContentType(ContentTypeEnum.BOOLEAN_DROP_DOWN_LIST);
            return;
        }
        if (td.getContent() != null && fieldType == LinkUrl.class) {
            td.setTdContentType(ContentTypeEnum.LINK_URL);
            setLinkTd(td);
            return;
        }
        if (td.getContent() != null && fieldType == LinkEmail.class) {
            td.setTdContentType(ContentTypeEnum.LINK_EMAIL);
            setLinkTd(td);
        }
    }

    private void setLinkTd(Td td) {
        String[] splits = td.getContent().split("->");
        if (splits.length == 1) {
            td.setLink(td.getContent());
        } else {
            td.setContent(splits[0]);
            td.setLink(splits[1]);
        }
    }

    /**
     * 初始化单元格样式
     */
    private void initStyleMap() {
        if (noStyle) {
            commonTdStyle = evenTdStyle = linkCommonStyle = linkEvenStyle = Collections.emptyMap();
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

            linkCommonStyle = new HashMap<>(commonTdStyle);
            linkCommonStyle.put(FontStyle.FONT_COLOR, "blue");
            linkCommonStyle.put(FontStyle.TEXT_DECORATION, FontStyle.UNDERLINE);

            linkEvenStyle = new HashMap<>(linkCommonStyle);
            linkEvenStyle.putAll(evenTdStyle);
        }
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
        if (customWidthMap == null) {
            customWidthMap = new HashMap<>(sortedFields.size());
        }

        boolean needToAddTitle = Objects.isNull(this.titles);
        for (int i = 0, size = sortedFields.size(); i < size; i++) {
            Field field = sortedFields.get(i);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            if (excelColumn != null) {
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
                if (widths == null && excelColumn.width() > 0) {
                    customWidthMap.put(i, excelColumn.width());
                }
                if (!noStyle && excelColumn.style().length > 0) {
                    setCustomStyle(i, excelColumn);
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

    private void setCustomStyle(int i, ExcelColumn excelColumn) {
        String[] styles = excelColumn.style();
        for (String style : styles) {
            if (StringUtil.isBlank(style)) {
                throw new IllegalArgumentException("Illegal style");
            }
            String[] splits = style.split("->");
            if (splits.length == 1) {
                // 发现未设置样式归属，则设置为全局样式，清除其他样式
                customStyle.put("cell&" + i, StyleUtil.parseStyle(splits[0]));
                break;
            } else {
                customStyle.put(splits[0] + "&" + i, StyleUtil.parseStyle(splits[1]));
            }
        }
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
        if (excelColumn == null) {
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
        if (excelColumn1 == null && excelColumn2 == null) {
            return 0;
        }
        int defaultOrder = 0;
        int order1 = defaultOrder;
        if (excelColumn1 != null) {
            order1 = excelColumn1.order();
        }
        int order2 = defaultOrder;
        if (excelColumn2 != null) {
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
        if (workbookType == null) {
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
        if (titles == null || titles.isEmpty()) {
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
     * @param <T>          泛型
     * @return 结果集
     */
    private <T> List<Pair<? extends Class, ?>> getRenderContent(T data, List<Field> sortedFields) {
        return sortedFields.stream()
                .map(field -> {
                    Pair<? extends Class, Object> value = WriteConverterContext.convert(field, data);
                    if (value.getValue() != null) {
                        return value;
                    }
                    String defaultValue = defaultValueMap.get(field);
                    if (defaultValue != null) {
                        return Pair.of(field.getType(), defaultValue);
                    }
                    if (globalDefaultValue != null) {
                        return Pair.of(field.getType(), globalDefaultValue);
                    }
                    return value;
                })
                .collect(Collectors.toCollection(LinkedList::new));
    }

}
