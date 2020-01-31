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
import com.github.liaochong.myexcel.core.annotation.ExcelModel;
import com.github.liaochong.myexcel.core.annotation.ExcelTable;
import com.github.liaochong.myexcel.core.annotation.ExcludeColumn;
import com.github.liaochong.myexcel.core.annotation.IgnoreColumn;
import com.github.liaochong.myexcel.core.constant.AllConverter;
import com.github.liaochong.myexcel.core.constant.BooleanDropDownList;
import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.constant.CsvConverter;
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

import javax.lang.model.type.NullType;
import java.io.File;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author liaochong
 * @version 1.0
 */
abstract class AbstractSimpleExcelBuilder {
    /**
     * 字段展示顺序
     */
    protected List<String> fieldDisplayOrder;
    /**
     * 已排序字段
     */
    protected List<Field> filteredFields = Collections.emptyList();
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
     * 默认值集合
     */
    protected Map<Field, String> defaultValueMap = new HashMap<>();
    /**
     * 自定义宽度
     */
    protected Map<Integer, Integer> customWidthMap = new HashMap<>();
    /**
     * 标题层级
     */
    protected int titleLevel = 0;
    /**
     * 自定义样式
     */
    protected Map<String, Map<String, String>> customStyle = new HashMap<>();
    /**
     * 是否为奇数行
     */
    protected boolean isOddRow = true;
    /**
     * 无样式
     */
    protected boolean noStyle;

    protected Map<Integer, Integer> widths;
    /**
     * 格式化
     */
    private Map<Integer, String> formats = new HashMap<>();
    /**
     * 格式样式Map
     */
    private Map<String, Map<String, String>> formatsStyleMap = new HashMap<>();
    /**
     * 是否为Map类型导出
     */
    protected boolean isMapBuild;
    /**
     * 全局设置
     */
    protected GlobalSetting globalSetting = new GlobalSetting();
    /**
     * ExcelColumn映射
     */
    private Map<Field, ExcelColumnMapping> excelColumnMappingMap = new HashMap<>();
    /**
     * 转换上下文
     */
    private ConvertContext convertContext = new ConvertContext(globalSetting, excelColumnMappingMap);

    /**
     * Core methods for obtaining export related fields, styles, etc
     *
     * @param classFieldContainer classFieldContainer
     * @param groups              分组
     * @return Field
     */
    protected List<Field> getFilteredFields(ClassFieldContainer classFieldContainer, Class<?>... groups) {
        setGlobalSetting(classFieldContainer);

        List<Field> preElectionFields = this.getPreElectionFields(classFieldContainer);
        List<Field> buildFields = this.getGroupFields(preElectionFields, groups);
        // 初始化标题容器
        List<String> titles = new ArrayList<>(buildFields.size());

        Map<String, String> globalStyleMap = getGlobalStyleMap(globalSetting.getGlobalStyle());
        this.setOddEvenStyle(globalStyleMap);
        for (int i = 0, size = buildFields.size(); i < size; i++) {
            Field field = buildFields.get(i);
            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
            setCustomStyle(field, i, globalStyleMap.get("cell"));
            setCustomStyle(field, i, globalStyleMap.get("title"));
            if (excelColumn != null) {
                if (globalSetting.isUseFieldNameAsTitle() && excelColumn.title().isEmpty()) {
                    titles.add(field.getName());
                } else {
                    titles.add(excelColumn.title());
                }
                if (!excelColumn.defaultValue().isEmpty()) {
                    defaultValueMap.put(field, excelColumn.defaultValue());
                }
                if (widths == null && excelColumn.width() > 0) {
                    customWidthMap.put(i, excelColumn.width());
                }
                if (excelColumn.style().length > 0) {
                    setCustomStyle(field, i, excelColumn.style());
                }
                if (!excelColumn.format().isEmpty()) {
                    formats.put(i, excelColumn.format());
                } else if (!excelColumn.decimalFormat().isEmpty()) {
                    formats.put(i, excelColumn.decimalFormat());
                } else if (!excelColumn.dateFormatPattern().isEmpty()) {
                    formats.put(i, excelColumn.dateFormatPattern());
                } else if (field.getType() == LocalDate.class) {
                    formats.put(i, globalSetting.getDateFormat());
                } else if (ReflectUtil.isDate(field.getType())) {
                    formats.put(i, globalSetting.getDateTimeFormat());
                } else if (ReflectUtil.isNumber(field.getType())) {
                    if (globalSetting.getDecimalFormat() != null) {
                        formats.put(i, globalSetting.getDecimalFormat());
                    }
                }
                ExcelColumnMapping mapping = ExcelColumnMapping.mapping(excelColumn);
                excelColumnMappingMap.put(field, mapping);
            } else {
                if (globalSetting.isUseFieldNameAsTitle()) {
                    titles.add(field.getName());
                } else {
                    titles.add(null);
                }
            }
        }
        setTitles(titles);
        if (!customWidthMap.isEmpty()) {
            globalSetting.setWidthStrategy(WidthStrategy.CUSTOM_WIDTH);
        }
        return buildFields;
    }

    /**
     * 创建table
     *
     * @return table
     */
    protected Table createTable() {
        Table table = new Table();
        table.setCaption(globalSetting.getSheetName());
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
            String[] multiTitles = title.split(globalSetting.getTitleSeparator());
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

        Map<String, String> thStyle = getDefaultThStyle();
        Map<Integer, List<Td>> rowTds = tdLists.stream().flatMap(List::stream).filter(td -> td.getRow() > -1).collect(Collectors.groupingBy(Td::getRow));
        List<Tr> trs = new ArrayList<>();
        boolean isComputeAutoWidth = WidthStrategy.isComputeAutoWidth(globalSetting.getWidthStrategy());
        rowTds.forEach((k, v) -> {
            Tr tr = new Tr(k, globalSetting.getTitleRowHeight());
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

    private Map<String, String> getDefaultThStyle() {
        if (!noStyle && customStyle.isEmpty()) {
            Map<String, String> thStyle = new HashMap<>(7);
            thStyle.put(FontStyle.FONT_WEIGHT, FontStyle.BOLD);
            thStyle.put(FontStyle.FONT_SIZE, "14");
            thStyle.put(TextAlignStyle.TEXT_ALIGN, TextAlignStyle.CENTER);
            thStyle.put(TextAlignStyle.VERTICAL_ALIGN, TextAlignStyle.MIDDLE);
            thStyle.put(BorderStyle.BORDER_BOTTOM_STYLE, BorderStyle.THIN);
            thStyle.put(BorderStyle.BORDER_LEFT_STYLE, BorderStyle.THIN);
            thStyle.put(BorderStyle.BORDER_RIGHT_STYLE, BorderStyle.THIN);
            return thStyle;
        }
        return Collections.emptyMap();
    }

    /**
     * 创建内容行
     *
     * @param contents 内容集合
     * @return 内容行
     */
    protected Tr createTr(List<Pair<? extends Class, ?>> contents) {
        Tr tr = new Tr(0, globalSetting.getRowHeight());
        if (contents.isEmpty()) {
            return tr;
        }
        boolean isComputeAutoWidth = WidthStrategy.isComputeAutoWidth(globalSetting.getWidthStrategy());
        boolean isCustomWidth = WidthStrategy.isCustomWidth(globalSetting.getWidthStrategy());
        tr.setColWidthMap((isComputeAutoWidth || isCustomWidth) ? new HashMap<>(contents.size()) : Collections.emptyMap());
        Map<String, String> tdStyle = isOddRow ? commonTdStyle : evenTdStyle;
        Map<String, String> linkStyle = isOddRow ? linkCommonStyle : linkEvenStyle;
        String oddEvenPrefix = isOddRow ? "odd&" : "even&";
        isOddRow = !isOddRow;
        boolean useCustomStyle = !noStyle && !customStyle.isEmpty();
        List<Td> tdList = IntStream.range(0, contents.size()).mapToObj(i -> {
            Td td = new Td(0, i);
            Pair<? extends Class, ?> pair = contents.get(i);
            setTdContent(td, pair);
            setTdContentType(td, pair.getKey());

            this.setFormula(i, td);

            Map<String, String> style;
            if (useCustomStyle) {
                style = customStyle.get(oddEvenPrefix + i);
                if (style == null) {
                    style = customStyle.getOrDefault("cell&" + i, Collections.emptyMap());
                }
            } else {
                style = ContentTypeEnum.isLink(td.getTdContentType()) ? linkStyle : tdStyle;
            }
            if (isComputeAutoWidth) {
                tr.getColWidthMap().put(i, TdUtil.getStringWidth(td.getContent()));
            }
            if (formats.get(i) != null) {
                String format = formats.get(i);
                Map<String, String> formatStyle = formatsStyleMap.get(format + "_" + i + "_" + oddEvenPrefix);
                if (formatStyle == null) {
                    formatStyle = new HashMap<>(style);
                    formatStyle.put("format", format);
                    formatsStyleMap.put(format + "_" + i, formatStyle);
                }
                style = formatStyle;
            }
            td.setStyle(style);
            return td;
        }).collect(Collectors.toList());
        if (isCustomWidth) {
            tr.setColWidthMap(customWidthMap);
        }
        tr.setTdList(tdList);
        return tr;
    }

    private void setFormula(int i, Td td) {
        if (filteredFields.isEmpty()) {
            return;
        }
        Field field = filteredFields.get(i);
        ExcelColumnMapping excelColumnMapping = excelColumnMappingMap.get(field);
        if (excelColumnMapping != null && excelColumnMapping.isFormula()) {
            td.setFormula(true);
        }
    }

    private void setTdContent(Td td, Pair<? extends Class, ?> pair) {
        Class fieldType = pair.getKey();
        if (fieldType == NullType.class) {
            return;
        }
        if (fieldType == Date.class) {
            td.setDate((Date) pair.getValue());
        } else if (fieldType == LocalDateTime.class) {
            td.setLocalDateTime((LocalDateTime) pair.getValue());
        } else if (fieldType == LocalDate.class) {
            td.setLocalDate((LocalDate) pair.getValue());
        } else if (com.github.liaochong.myexcel.core.constant.File.class.isAssignableFrom(fieldType)) {
            td.setFile((File) pair.getValue());
        } else {
            td.setContent(String.valueOf(pair.getValue()));
        }
    }

    private void setTdContentType(Td td, Class fieldType) {
        if (String.class == fieldType) {
            return;
        }
        if (ReflectUtil.isNumber(fieldType)) {
            td.setTdContentType(ContentTypeEnum.DOUBLE);
            return;
        }
        if (ReflectUtil.isDate(fieldType)) {
            td.setTdContentType(ContentTypeEnum.DATE);
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
        String[] splits = td.getContent().split(Constants.ARROW);
        if (splits.length == 1) {
            td.setLink(td.getContent());
        } else {
            td.setContent(splits[0]);
            td.setLink(splits[1]);
        }
    }

    protected void setTitles(List<String> titles) {
        if (this.titles == null) {
            boolean hasTitle = titles.stream().anyMatch(StringUtil::isNotBlank);
            if (hasTitle) {
                this.titles = titles;
            }
        }
    }

    protected List<Field> getGroupFields(List<Field> preElectionFields, Class<?>[] groups) {
        List<Class<?>> selectedGroupList = Objects.nonNull(groups) ? Arrays.stream(groups).filter(Objects::nonNull).collect(Collectors.toList()) : Collections.emptyList();
        return preElectionFields.stream()
                .filter(field -> (!field.isAnnotationPresent(ExcludeColumn.class) && !field.isAnnotationPresent(IgnoreColumn.class)) && ReflectUtil.isFieldSelected(selectedGroupList, field))
                .sorted(ReflectUtil::sortFields)
                .collect(Collectors.toList());
    }


    protected void setGlobalSetting(ClassFieldContainer classFieldContainer) {
        ClassFieldContainer parentContainer = classFieldContainer.getParent();
        if (parentContainer != null) {
            setGlobalSetting(parentContainer);
        }
        if (classFieldContainer.getClazz() == Object.class) {
            return;
        }
        ExcelModel excelModel = classFieldContainer.getClazz().getAnnotation(ExcelModel.class);
        if (excelModel == null) {
            ExcelTable excelTable = classFieldContainer.getClazz().getAnnotation(ExcelTable.class);
            if (excelTable == null) {
                return;
            }
            if (!excelTable.sheetName().isEmpty() && globalSetting.getSheetName() == null) {
                globalSetting.setSheetName(excelTable.sheetName());
            }
            if (excelTable.workbookType() != WorkbookType.SXLSX && globalSetting.getWorkbookType() == null) {
                globalSetting.setWorkbookType(excelTable.workbookType());
            }
            if (excelTable.excludeParent()) {
                globalSetting.setExcludeParent(true);
            }
            if (!excelTable.includeAllField()) {
                globalSetting.setIncludeAllField(false);
            }
            if (!excelTable.defaultValue().isEmpty()) {
                globalSetting.setDefaultValue(excelTable.defaultValue());
            }
            if (!excelTable.wrapText()) {
                globalSetting.setWrapText(false);
            }
            if (!excelTable.titleSeparator().equals(Constants.ARROW)) {
                globalSetting.setTitleSeparator(excelTable.titleSeparator());
            }
            if (!excelTable.ignoreStaticFields()) {
                globalSetting.setIgnoreStaticFields(false);
            }
            if (excelTable.titleRowHeight() != -1) {
                globalSetting.setTitleRowHeight(excelTable.titleRowHeight());
            }
            if (excelTable.rowHeight() != -1) {
                globalSetting.setRowHeight(excelTable.rowHeight());
            }
            if (excelTable.style().length != 0 && globalSetting.getGlobalStyle().isEmpty()) {
                globalSetting.getGlobalStyle().addAll(Arrays.asList(excelTable.style()));
            }
            if (excelTable.useFieldNameAsTitle()) {
                globalSetting.setUseFieldNameAsTitle(true);
            }
        } else {
            if (!excelModel.sheetName().isEmpty() && globalSetting.getSheetName() == null) {
                globalSetting.setSheetName(excelModel.sheetName());
            }
            if (excelModel.workbookType() != WorkbookType.SXLSX && globalSetting.getWorkbookType() == null) {
                globalSetting.setWorkbookType(excelModel.workbookType());
            }
            if (excelModel.excludeParent()) {
                globalSetting.setExcludeParent(true);
            }
            if (!excelModel.includeAllField()) {
                globalSetting.setIncludeAllField(false);
            }
            if (!excelModel.defaultValue().isEmpty()) {
                globalSetting.setDefaultValue(excelModel.defaultValue());
            }
            if (!excelModel.wrapText()) {
                globalSetting.setWrapText(false);
            }
            if (!excelModel.titleSeparator().equals(Constants.ARROW)) {
                globalSetting.setTitleSeparator(excelModel.titleSeparator());
            }
            if (!excelModel.ignoreStaticFields()) {
                globalSetting.setIgnoreStaticFields(false);
            }
            if (excelModel.titleRowHeight() != -1) {
                globalSetting.setTitleRowHeight(excelModel.titleRowHeight());
            }
            if (excelModel.rowHeight() != -1) {
                globalSetting.setRowHeight(excelModel.rowHeight());
            }
            if (excelModel.style().length != 0 && globalSetting.getGlobalStyle().isEmpty()) {
                globalSetting.getGlobalStyle().addAll(Arrays.asList(excelModel.style()));
            }
            if (excelModel.useFieldNameAsTitle()) {
                globalSetting.setUseFieldNameAsTitle(true);
            }
            if (!excelModel.decimalFormat().isEmpty()) {
                globalSetting.setDecimalFormat(excelModel.decimalFormat());
            }
            if (!excelModel.dateFormat().isEmpty()) {
                globalSetting.setDateFormat(excelModel.dateFormat());
            }
            if (!excelModel.dateTimeFormat().isEmpty()) {
                globalSetting.setDateTimeFormat(excelModel.dateTimeFormat());
            }
        }
    }

    private Map<String, String> getGlobalStyleMap(Set<String> globalStyle) {
        Map<String, String> globalStyleMap = new HashMap<>();
        if (globalStyle != null) {
            globalStyle.forEach(style -> {
                String[] splits = style.split(Constants.ARROW);
                if (splits.length == 1) {
                    globalStyleMap.put("cell", style);
                } else {
                    globalStyleMap.put(splits[0], style);
                }
            });
        }
        return globalStyleMap;
    }

    private void setOddEvenStyle(Map<String, String> globalStyleMap) {
        String oddStyle = globalStyleMap.get("odd");
        if (oddStyle != null) {
            commonTdStyle = StyleUtil.parseStyle(oddStyle.split(Constants.ARROW)[1]);
        }
        String evenStyle = globalStyleMap.get("even");
        if (evenStyle != null) {
            evenTdStyle = StyleUtil.parseStyle(evenStyle.split(Constants.ARROW)[1]);
        }
    }

    private void setCustomStyle(Field field, int index, String... styles) {
        for (String style : styles) {
            if (StringUtil.isBlank(style)) {
                continue;
            }
            if (StringUtil.isBlank(style)) {
                throw new IllegalArgumentException("Illegal style,field:" + field.getName());
            }
            String[] splits = style.split(Constants.ARROW);
            if (splits.length == 1) {
                // 发现未设置样式归属，则设置为全局样式，清除其他样式
                Map<String, String> styleMap = setWidthStrategyAndWidth(splits, 0, index);
                customStyle.put("cell&" + index, styleMap);
                break;
            } else {
                Map<String, String> styleMap = setWidthStrategyAndWidth(splits, 1, index);
                customStyle.put(splits[0] + "&" + index, styleMap);
            }
        }
    }

    private Map<String, String> setWidthStrategyAndWidth(String[] splits, int splitIndex, int fieldIndex) {
        Map<String, String> styleMap = StyleUtil.parseStyle(splits[splitIndex]);
        String width = styleMap.get("width");
        if (width != null) {
            globalSetting.setWidthStrategy(WidthStrategy.CUSTOM_WIDTH);
            customWidthMap.put(fieldIndex, TdUtil.getValue(width));
        }
        return styleMap;
    }

    protected List<Field> getPreElectionFields(ClassFieldContainer classFieldContainer) {
        if (Objects.nonNull(fieldDisplayOrder) && !fieldDisplayOrder.isEmpty()) {
            this.selfAdaption();
            return fieldDisplayOrder.stream()
                    .map(classFieldContainer::getFieldByName)
                    .collect(Collectors.toList());
        }
        List<Field> preElectionFields;
        if (globalSetting.isIncludeAllField()) {
            if (globalSetting.isExcludeParent()) {
                preElectionFields = classFieldContainer.getDeclaredFields();
            } else {
                preElectionFields = classFieldContainer.getFields();
            }
        } else {
            if (globalSetting.isExcludeParent()) {
                preElectionFields = classFieldContainer.getDeclaredFields().stream()
                        .filter(field -> field.isAnnotationPresent(ExcelColumn.class))
                        .collect(Collectors.toList());
            } else {
                preElectionFields = classFieldContainer.getFieldsByAnnotation(ExcelColumn.class);
            }
        }
        if (globalSetting.isIgnoreStaticFields()) {
            preElectionFields = preElectionFields.stream()
                    .filter(field -> !Modifier.isStatic(field.getModifiers()))
                    .collect(Collectors.toList());
        }
        return preElectionFields;
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
     * @param csv          是否是csv构建
     * @param <T>          泛型
     * @return 结果集
     */
    protected <T> List<Pair<? extends Class, ?>> getRenderContent(T data, List<Field> sortedFields, boolean csv) {
        convertContext.setConverterType(csv ? CsvConverter.class : AllConverter.class);
        return sortedFields.stream()
                .map(field -> {
                    Pair<? extends Class, Object> value = WriteConverterContext.convert(field, data, convertContext);
                    if (value.getValue() != null) {
                        return value;
                    }
                    String defaultValue = defaultValueMap.get(field);
                    if (defaultValue != null) {
                        return Pair.of(field.getType(), defaultValue);
                    }
                    if (globalSetting.getDefaultValue() != null) {
                        return Pair.of(field.getType(), globalSetting.getDefaultValue());
                    }
                    return value;
                })
                .collect(Collectors.toCollection(LinkedList::new));
    }

    protected List<Pair<? extends Class, ?>> assemblingMapContents(Map<String, Object> data) {
        if (data == null || data.isEmpty()) {
            return Collections.emptyList();
        }
        List<Pair<? extends Class, ?>> contents = new ArrayList<>(data.size());
        if (fieldDisplayOrder == null) {
            data.forEach((k, v) -> {
                contents.add(Pair.of(v == null ? String.class : v.getClass(), v));
            });
        } else {
            for (String fieldName : fieldDisplayOrder) {
                Object val = data.get(fieldName);
                contents.add(Pair.of(val == null ? String.class : val.getClass(), val));
            }
        }
        return contents;
    }

    /**
     * 初始化单元格样式
     */
    protected void initStyleMap() {
        if (!noStyle && customStyle.isEmpty()) {
            if (commonTdStyle == null) {
                commonTdStyle = new HashMap<>(3);
                commonTdStyle.put(BorderStyle.BORDER_BOTTOM_STYLE, BorderStyle.THIN);
                commonTdStyle.put(BorderStyle.BORDER_LEFT_STYLE, BorderStyle.THIN);
                commonTdStyle.put(BorderStyle.BORDER_RIGHT_STYLE, BorderStyle.THIN);
                commonTdStyle.put(TextAlignStyle.VERTICAL_ALIGN, TextAlignStyle.MIDDLE);
                if (globalSetting.isWrapText()) {
                    commonTdStyle.put(WordBreakStyle.WORD_BREAK, WordBreakStyle.BREAK_ALL);
                }
            }
            if (evenTdStyle == null) {
                evenTdStyle = new HashMap<>(4);
                evenTdStyle.put(BackgroundStyle.BACKGROUND_COLOR, "#f6f8fa");
                evenTdStyle.putAll(commonTdStyle);
            }
            linkCommonStyle = new HashMap<>(commonTdStyle);
            linkCommonStyle.put(FontStyle.FONT_COLOR, "blue");
            linkCommonStyle.put(FontStyle.TEXT_DECORATION, FontStyle.UNDERLINE);

            linkEvenStyle = new HashMap<>(linkCommonStyle);
            linkEvenStyle.putAll(evenTdStyle);
        } else {
            commonTdStyle = evenTdStyle = linkCommonStyle = linkEvenStyle = Collections.emptyMap();
        }
    }
}
