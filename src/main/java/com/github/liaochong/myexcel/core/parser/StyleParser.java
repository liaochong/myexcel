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
package com.github.liaochong.myexcel.core.parser;

import com.github.liaochong.myexcel.core.constant.Constants;
import com.github.liaochong.myexcel.core.style.FontStyle;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.StyleUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import lombok.Setter;

import java.lang.reflect.Field;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * 单元格所有样式解析以及提供
 * <p>
 * title样式仅支持title
 * <p>
 * cell样式规则，继承关系：globalCommonStyle(globalEvenStyle)->globalLinkStyle(globalLinkEvenStyle)->globalCellStyle,
 * 如存在相同样式，子样式覆盖父样式
 *
 * @author liaochong
 * @version 1.0
 */
public final class StyleParser {
    /**
     * 全局标题样式
     */
    private Map<String, String> globalTitleStyle = Collections.emptyMap();
    /**
     * 全局单元格样式
     */
    private Map<String, String> globalCommonStyle = Collections.emptyMap();
    /**
     * 全局奇数行单元格样式
     */
    private Map<String, String> globalEvenStyle = Collections.emptyMap();
    /**
     * 全局超链接样式
     */
    private Map<String, String> globalLinkStyle;
    /**
     * 全局超链接奇数行样式
     */
    private Map<String, String> globalLinkEvenStyle;
    /**
     * 全局单元格样式
     */
    private Map<String, String> globalCellStyle;
    /**
     * 各个列样式
     */
    private Map<String, Map<String, String>> eachColumnStyle = new HashMap<>();
    /**
     * 格式样式Map
     */
    private Map<String, Map<String, String>> formatsStyleMap = new HashMap<>();
    /**
     * 自定义宽度
     */
    @Setter
    private Map<Integer, Integer> customWidthMap;
    /**
     * 是否为偶数行
     */
    private boolean isOddRow = true;
    /**
     * 无样式
     */
    @Setter
    private boolean noStyle;

    public StyleParser(Map<Integer, Integer> customWidthMap) {
        this.customWidthMap = customWidthMap;
    }

    /**
     * 解析全局样式，{@link com.github.liaochong.myexcel.core.annotation.ExcelModel}
     *
     * @param styles 全局样式
     */
    public void parse(Set<String> styles) {
        if (noStyle) {
            return;
        }
        Map<String, String> styleMap = new HashMap<>();
        styles.forEach(style -> {
            String[] splits = style.split(Constants.ARROW);
            if (splits.length == 1) {
                styleMap.putIfAbsent("cell", style);
                return;
            }
            boolean appoint = splits[0].contains("&");
            if (appoint) {
                eachColumnStyle.put(splits[0], StyleUtil.parseStyle(splits[1]));
            } else {
                styleMap.putIfAbsent(splits[0], style);
            }
        });
        String titleStyle = styleMap.get("title");
        if (titleStyle != null) {
            globalTitleStyle = StyleUtil.parseStyle(titleStyle.split(Constants.ARROW)[1]);
        }
        String oddStyle = styleMap.get("odd");
        if (oddStyle != null) {
            globalCommonStyle = StyleUtil.parseStyle(oddStyle.split(Constants.ARROW)[1]);
        }
        String evenStyle = styleMap.get("even");
        if (evenStyle != null) {
            globalEvenStyle = StyleUtil.parseStyle(evenStyle.split(Constants.ARROW)[1]);
        }
        globalLinkStyle = new HashMap<>(globalCommonStyle);
        globalLinkStyle.put(FontStyle.FONT_COLOR, "blue");
        globalLinkStyle.put(FontStyle.TEXT_DECORATION, FontStyle.UNDERLINE);

        globalLinkEvenStyle = new HashMap<>(globalLinkStyle);
        globalLinkEvenStyle.putAll(globalEvenStyle);

        String cellStyle = styleMap.get("cell");
        if (cellStyle != null) {
            this.globalCellStyle = StyleUtil.parseStyle(cellStyle.split(Constants.ARROW)[1]);
        }
    }

    public void setColumnStyle(Field field, int fieldIndex, String... columnStyles) {
        if (noStyle) {
            return;
        }
        setEachColumnStyle("title", fieldIndex, globalTitleStyle);
        setEachColumnStyle("even", fieldIndex, globalEvenStyle);
        setEachColumnStyle("odd", fieldIndex, globalCommonStyle);
        setEachColumnStyle("cell", fieldIndex, globalCellStyle);
        if (columnStyles == null) {
            return;
        }
        for (String columnStyle : columnStyles) {
            if (StringUtil.isBlank(columnStyle)) {
                throw new IllegalArgumentException("Illegal style,field:" + field.getName());
            }
            String[] splits = columnStyle.split(Constants.ARROW);
            if (splits.length == 1) {
                // 发现未设置样式归属，则设置为全局样式，清除其他样式
                setEachColumnStyle("cell", fieldIndex, StyleUtil.parseStyle(splits[0]));
            } else {
                setEachColumnStyle(splits[0], fieldIndex, StyleUtil.parseStyle(splits[1]));
            }
        }
    }

    private void setEachColumnStyle(String prefix, int fieldIndex, Map<String, String> styleMap) {
        if (styleMap == null || styleMap.isEmpty()) {
            return;
        }
        String stylePrefix = prefix + "&" + fieldIndex;
        Map<String, String> parentStyleMap = eachColumnStyle.get(stylePrefix);
        if (parentStyleMap == null || parentStyleMap.isEmpty()) {
            parentStyleMap = new HashMap<>(styleMap);
            eachColumnStyle.put(stylePrefix, parentStyleMap);
        } else {
            parentStyleMap.putAll(styleMap);
        }
        setWidth(fieldIndex, parentStyleMap);
    }

    public Map<String, String> getTitleStyle(String styleKey) {
        if (noStyle) {
            return Collections.emptyMap();
        }
        return eachColumnStyle.getOrDefault(styleKey, globalTitleStyle);
    }

    public void toggle() {
        isOddRow = !isOddRow;
    }

    public Map<String, String> getCellStyle(int fieldIndex, ContentTypeEnum contentType, String format) {
        Map<String, String> style = Collections.emptyMap();
        if (!noStyle) {
            Map<String, String> tdStyle = isOddRow ? globalCommonStyle : globalEvenStyle;
            Map<String, String> linkStyle = isOddRow ? globalLinkStyle : globalLinkEvenStyle;
            String oddEvenPrefix = isOddRow ? "odd&" : "even&";
            style = eachColumnStyle.get(oddEvenPrefix + fieldIndex);
            Map<String, String> cellStyleMap = eachColumnStyle.get("cell&" + fieldIndex);
            if (cellStyleMap != null) {
                if (style == null || style.isEmpty()) {
                    style = cellStyleMap;
                } else {
                    style = new HashMap<>(style);
                    style.putAll(cellStyleMap);
                }
            }
            if (style == null) {
                style = ContentTypeEnum.isLink(contentType) ? linkStyle : tdStyle;
            }
        }
        if (format != null) {
            style = this.getFormatStyle(format, fieldIndex, style);
        }
        return style;
    }

    private Map<String, String> getFormatStyle(String format, int fieldIndex, Map<String, String> style) {
        Map<String, String> formatStyle = formatsStyleMap.get(format + "_" + fieldIndex + "_" + (isOddRow ? "odd&" : "even&"));
        if (formatStyle == null) {
            formatStyle = new HashMap<>(style);
            formatStyle.put("format", format);
            formatsStyleMap.put(format + "_" + fieldIndex, formatStyle);
        }
        return formatStyle;
    }

    private void setWidth(int fieldIndex, Map<String, String> styleMap) {
        String width = styleMap.get("width");
        if (width != null) {
            customWidthMap.put(fieldIndex, TdUtil.getValue(width));
        }
    }
}
