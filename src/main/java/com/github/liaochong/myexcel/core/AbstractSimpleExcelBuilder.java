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

import com.github.liaochong.myexcel.core.annotation.ExcelColumn;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import com.github.liaochong.myexcel.core.annotation.ExcludeColumn;
import com.github.liaochong.myexcel.core.constant.BooleanDropDownList;
import com.github.liaochong.myexcel.core.constant.DropDownList;
import com.github.liaochong.myexcel.core.constant.ImageFile;
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
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;
import com.github.liaochong.myexcel.core.style.BackgroundStyle;
import com.github.liaochong.myexcel.core.style.BorderStyle;
import com.github.liaochong.myexcel.core.style.FontStyle;
import com.github.liaochong.myexcel.core.style.TextAlignStyle;
import com.github.liaochong.myexcel.core.style.WordBreakStyle;
import com.github.liaochong.myexcel.utils.ReflectUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.StyleUtil;
import com.github.liaochong.myexcel.utils.TdUtil;

import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
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
abstract class AbstractSimpleExcelBuilder implements SimpleStreamExcelBuilder {
    /**
     * 字段展示顺序
     */
    protected List<String> fieldDisplayOrder;
    /**
     * excel workbook
     */
    protected WorkbookType workbookType;
    /**
     * 一般单元格样式
     */
    protected Map<String, String> commonTdStyle;
    /**
     * 偶数行单元格样式
     */
    protected Map<String, String> evenTdStyle;
    /**
     * 超链接公共样式
     */
    protected Map<String, String> linkCommonStyle;
    /**
     * 超链接偶数行样式
     */
    protected Map<String, String> linkEvenStyle;
    /**
     * 标题
     */
    protected List<String> titles;
    /**
     * sheetName
     */
    protected String sheetName;
    /**
     * 全局默认值
     */
    protected String globalDefaultValue;
    /**
     * 默认值集合
     */
    protected Map<Field, String> defaultValueMap;
    /**
     * 自定义宽度
     */
    protected Map<Integer, Integer> customWidthMap;
    /**
     * 是否自动换行
     */
    protected boolean wrapText = true;
    /**
     * 标题层级
     */
    protected int titleLevel = 0;
    /**
     * 标题分离器
     */
    protected String titleSeparator = "->";
    /**
     * 自定义样式
     */
    protected Map<String, Map<String, String>> customStyle = new HashMap<>();
    /**
     * 是否为奇数行
     */
    protected boolean isOddRow = true;
    /**
     * 行高
     */
    protected int rowHeight;
    /**
     * 标题行行高
     */
    protected int titleRowHeight;
    /**
     * 无样式
     */
    protected boolean noStyle;
    /**
     * 自动宽度策略
     */
    protected WidthStrategy widthStrategy;

    protected Map<Integer, Integer> widths;

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
        boolean isComputeAutoWidth = WidthStrategy.isComputeAutoWidth(widthStrategy);
        rowTds.forEach((k, v) -> {
            Tr tr = new Tr(k, titleRowHeight);
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
    protected Tr createTr(List<Pair<? extends Class, ?>> contents) {
        boolean isComputeAutoWidth = WidthStrategy.isComputeAutoWidth(widthStrategy);
        boolean isCustomWidth = WidthStrategy.isCustomWidth(widthStrategy);
        Tr tr = new Tr(0, rowHeight);
        tr.setColWidthMap((isComputeAutoWidth || isCustomWidth) ? new HashMap<>(contents.size()) : Collections.emptyMap());
        Map<String, String> tdStyle = isOddRow ? commonTdStyle : evenTdStyle;
        Map<String, String> linkStyle = isOddRow ? linkCommonStyle : linkEvenStyle;
        isOddRow = !isOddRow;
        List<Td> tdList = IntStream.range(0, contents.size()).mapToObj(i -> {
            Td td = new Td(0, i);

            Pair<? extends Class, ?> pair = contents.get(i);
            Class fieldType = pair.getKey();
            if (com.github.liaochong.myexcel.core.constant.File.class.isAssignableFrom(fieldType)) {
                td.setFile(pair.getValue() == null ? null : (File) pair.getValue());
            } else {
                td.setContent(pair.getValue() == null ? null : String.valueOf(pair.getValue()));
            }
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
            return;
        }
        if (td.getFile() != null && fieldType == ImageFile.class) {
            td.setTdContentType(ContentTypeEnum.IMAGE);
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
    protected void initStyleMap() {
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
            titleRowHeight = excelTable.titleRowHeight();
            rowHeight = excelTable.rowHeight();
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
                .filter(field -> !field.isAnnotationPresent(ExcludeColumn.class) && ReflectUtil.isFieldSelected(selectedGroupList, field))
                .sorted(ReflectUtil::sortFields)
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

    /**
     * 设置workbook
     *
     * @param excelTable excelTable
     */
    private void setWorkbookWithExcelTableAnnotation(ExcelTable excelTable) {
        if (workbookType == null) {
            this.workbookType = excelTable.workbookType();
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
    protected <T> List<Pair<? extends Class, ?>> getRenderContent(T data, List<Field> sortedFields) {
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
