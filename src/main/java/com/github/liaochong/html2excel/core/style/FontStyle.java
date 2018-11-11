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
package com.github.liaochong.html2excel.core.style;

import com.github.liaochong.html2excel.core.cache.FontCache;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public final class FontStyle {

    private static final short DEFAULT_FONT_SIZE = 12;

    public static void setFont(Workbook workbook, Row row, CellStyle cellStyle, Map<String, String> tdStyle, Map<Integer, Short> maxTdHeightMap) {
        String cacheKey = getCacheKey(tdStyle);
        if (FontCache.contains(cacheKey)) {
            cellStyle.setFont(FontCache.get(cacheKey));
            return;
        }
        Font font = null;
        String fs = tdStyle.get("font-size");
        if (Objects.nonNull(fs)) {
            fs = fs.replaceAll("\\D*", "");
            short fontSize = Short.parseShort(fs);
            font = workbook.createFont();
            font.setFontHeightInPoints(fontSize);
            if (fontSize > maxTdHeightMap.getOrDefault(row.getRowNum(), DEFAULT_FONT_SIZE)) {
                maxTdHeightMap.put(row.getRowNum(), fontSize);
            }
        }
        String fontFamily = tdStyle.get("font-family");
        if (Objects.nonNull(fontFamily)) {
            font = createFontIfNull(workbook, font);
            font.setFontName(fontFamily);
        }
        String italic = tdStyle.get("font-style");
        if (Objects.equals("italic", italic)) {
            font = createFontIfNull(workbook, font);
            font.setItalic(true);
        }
        String strikeout = tdStyle.get("text-decoration");
        if (Objects.equals(strikeout, "line-through")) {
            font = createFontIfNull(workbook, font);
            font.setStrikeout(true);
        }
        String fontWeight = tdStyle.get("font-weight");
        if (Objects.equals(fontWeight, "bold")) {
            font = createFontIfNull(workbook, font);
            font.setBold(true);
        }
        if (Objects.nonNull(font)) {
            cellStyle.setFont(font);
            FontCache.cache(cacheKey, font);
        }
    }

    private static Font createFontIfNull(Workbook workbook, Font font) {
        if (Objects.isNull(font)) {
            font = workbook.createFont();
        }
        return font;
    }

    private static String getCacheKey(Map<String, String> tdStyle) {
        StringBuilder result = new StringBuilder();
        appendKey(tdStyle, "font-size", result);
        appendKey(tdStyle, "font-family", result);
        appendKey(tdStyle, "font-style", result);
        appendKey(tdStyle, "text-decoration", result);
        appendKey(tdStyle, "font-weight", result);
        return result.toString();
    }

    private static void appendKey(Map<String, String> tdStyle, String styleName, StringBuilder result) {
        String style = tdStyle.get(styleName);
        if (Objects.nonNull(style)) {
            result.append(styleName).append(":").append(style).append("_");
        }
    }
}
