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
 * @author liaochong
 * @version 1.0
 */
public final class StyleParser {
    /**
     * 标题样式
     */
    private Map<String, String> titleTdStyle = Collections.emptyMap();
    /**
     * 一般单元格样式
     */
    private Map<String, String> commonTdStyle = Collections.emptyMap();
    /**
     * 偶数行单元格样式
     */
    private Map<String, String> evenTdStyle = Collections.emptyMap();
    /**
     * 超链接公共样式
     */
    private Map<String, String> linkCommonStyle;
    /**
     * 超链接偶数行样式
     */
    private Map<String, String> linkEvenStyle;

    private Map<String, String> cellStyle;
    /**
     * 自定义样式
     */
    private Map<String, Map<String, String>> customStyle = new HashMap<>();

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
     * 是否为奇数行
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
                customStyle.put(splits[0], StyleUtil.parseStyle(splits[1]));
            } else {
                styleMap.putIfAbsent(splits[0], style);
            }
        });
        String titleStyle = styleMap.get("title");
        if (titleStyle != null) {
            titleTdStyle = StyleUtil.parseStyle(titleStyle.split(Constants.ARROW)[1]);
        }
        String oddStyle = styleMap.get("odd");
        if (oddStyle != null) {
            commonTdStyle = StyleUtil.parseStyle(oddStyle.split(Constants.ARROW)[1]);
        }
        String evenStyle = styleMap.get("even");
        if (evenStyle != null) {
            evenTdStyle = StyleUtil.parseStyle(evenStyle.split(Constants.ARROW)[1]);
        }
        linkCommonStyle = new HashMap<>(commonTdStyle);
        linkCommonStyle.put(FontStyle.FONT_COLOR, "blue");
        linkCommonStyle.put(FontStyle.TEXT_DECORATION, FontStyle.UNDERLINE);

        linkEvenStyle = new HashMap<>(linkCommonStyle);
        linkEvenStyle.putAll(evenTdStyle);

        String cellStyle = styleMap.get("cell");
        if (cellStyle != null) {
            this.cellStyle = StyleUtil.parseStyle(cellStyle.split(Constants.ARROW)[1]);
        }
    }

    public void setColumnStyle(Field field, int fieldIndex, String... columnStyles) {
        if (noStyle) {
            return;
        }
        setCustomStyle("title", fieldIndex, titleTdStyle);
        setCustomStyle("even", fieldIndex, evenTdStyle);
        setCustomStyle("odd", fieldIndex, commonTdStyle);
        setCustomStyle("cell", fieldIndex, cellStyle);
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
                setCustomStyle("cell", fieldIndex, StyleUtil.parseStyle(splits[0]));
            } else {
                setCustomStyle(splits[0], fieldIndex, StyleUtil.parseStyle(splits[1]));
            }
        }
    }

    private void setCustomStyle(String prefix, int fieldIndex, Map<String, String> styleMap) {
        if (styleMap == null || styleMap.isEmpty()) {
            return;
        }
        String stylePrefix = prefix + "&" + fieldIndex;
        Map<String, String> parentStyleMap = customStyle.get(stylePrefix);
        if (parentStyleMap == null || parentStyleMap.isEmpty()) {
            parentStyleMap = new HashMap<>(styleMap);
            customStyle.put(stylePrefix, parentStyleMap);
        } else {
            parentStyleMap.putAll(styleMap);
        }
        setWidth(fieldIndex, parentStyleMap);
    }

    public Map<String, String> getTitleStyle(String styleKey) {
        if (noStyle) {
            return Collections.emptyMap();
        }
        return customStyle.getOrDefault(styleKey, titleTdStyle);
    }

    public void toggle() {
        isOddRow = !isOddRow;
    }

    public Map<String, String> getCellStyle(int fieldIndex, ContentTypeEnum contentType, String format) {
        Map<String, String> style = Collections.emptyMap();
        if (!noStyle) {
            Map<String, String> tdStyle = isOddRow ? commonTdStyle : evenTdStyle;
            Map<String, String> linkStyle = isOddRow ? linkCommonStyle : linkEvenStyle;
            String oddEvenPrefix = isOddRow ? "odd&" : "even&";
            style = customStyle.get(oddEvenPrefix + fieldIndex);
            Map<String, String> cellStyleMap = customStyle.get("cell&" + fieldIndex);
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
