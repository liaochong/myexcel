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

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public final class FontStyle {

    public static final String FONT_SIZE = "font-size";

    public static final String FONT_WEIGHT = "font-weight";

    public static final String FONT_FAMILY = "font-family";

    public static final String FONT_STYLE = "font-style";

    public static final String TEXT_DECORATION = "text-decoration";

    public static final String ITALIC = "italic";

    public static final String LINE_THROUGH = "line-through";

    public static final String BOLD = "bold";

    public static final short DEFAULT_FONT_SIZE = 12;

    public static final double FONT_SIZE_SHIFT = 0.25;

    public static double getFontWidthShift(Map<String, String> tdStyle) {
        double fontWidthShift = 0;
        String fontSize = tdStyle.get(FONT_SIZE);
        if (Objects.nonNull(fontSize)) {
            short fontSizeVal = Short.parseShort(fontSize.replaceAll("\\D*", ""));
            int intervalSize = fontSizeVal - FontStyle.DEFAULT_FONT_SIZE;
            if (intervalSize > 0) {
                fontWidthShift = intervalSize * FONT_SIZE_SHIFT;
            }
        }
        String fontWeight = tdStyle.get(FONT_WEIGHT);
        if (Objects.equals(fontWeight, BOLD)) {
            fontWidthShift += FONT_SIZE_SHIFT;
        }
        return fontWidthShift;
    }

    public static void setFont(Workbook workbook, CellStyle cellStyle, Map<String, String> tdStyle, Map<String, Font> fontMap) {
        String cacheKey = getCacheKey(tdStyle);
        if (Objects.nonNull(fontMap.get(cacheKey))) {
            cellStyle.setFont(fontMap.get(cacheKey));
            return;
        }
        Font font = null;
        String fs = tdStyle.get(FONT_SIZE);
        if (Objects.nonNull(fs)) {
            fs = fs.replaceAll("\\D*", "");
            short fontSize = Short.parseShort(fs);
            font = workbook.createFont();
            font.setFontHeightInPoints(fontSize);
        }
        String fontFamily = tdStyle.get(FONT_FAMILY);
        if (Objects.nonNull(fontFamily)) {
            font = createFontIfNull(workbook, font);
            font.setFontName(fontFamily);
        }
        String italic = tdStyle.get(FONT_STYLE);
        if (Objects.equals(ITALIC, italic)) {
            font = createFontIfNull(workbook, font);
            font.setItalic(true);
        }
        String strikeout = tdStyle.get(TEXT_DECORATION);
        if (Objects.equals(strikeout, LINE_THROUGH)) {
            font = createFontIfNull(workbook, font);
            font.setStrikeout(true);
        }
        String fontWeight = tdStyle.get(FONT_WEIGHT);
        if (Objects.equals(fontWeight, BOLD)) {
            font = createFontIfNull(workbook, font);
            font.setBold(true);
        }
        if (Objects.nonNull(font)) {
            cellStyle.setFont(font);
            fontMap.put(cacheKey, font);
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
        appendKey(tdStyle, FONT_SIZE, result);
        appendKey(tdStyle, FONT_FAMILY, result);
        appendKey(tdStyle, FONT_STYLE, result);
        appendKey(tdStyle, TEXT_DECORATION, result);
        appendKey(tdStyle, FONT_WEIGHT, result);
        return result.toString();
    }

    private static void appendKey(Map<String, String> tdStyle, String styleName, StringBuilder result) {
        String style = tdStyle.get(styleName);
        if (Objects.nonNull(style)) {
            result.append(styleName).append(":").append(style).append("_");
        }
    }
}
