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
package com.github.liaochong.myexcel.core.style;

import com.github.liaochong.myexcel.utils.ColorUtil;
import com.github.liaochong.myexcel.utils.StringUtil;
import com.github.liaochong.myexcel.utils.TdUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import java.awt.*;
import java.util.Map;
import java.util.function.Supplier;

/**
 * 字体样式
 *
 * @author liaochong
 * @version 1.0
 */
public final class FontStyle {

    public static final String FONT_COLOR = "color";

    public static final String FONT_SIZE = "font-size";

    public static final String FONT_WEIGHT = "font-weight";

    public static final String FONT_FAMILY = "font-family";

    public static final String FONT_STYLE = "font-style";

    public static final String TEXT_DECORATION = "text-decoration";

    public static final String ITALIC = "italic";

    public static final String LINE_THROUGH = "line-through";

    public static final String UNDERLINE = "underline";

    public static final String BOLD = "bold";

    public static final short DEFAULT_FONT_SIZE = 12;

    public static void setFont(Supplier<Font> fontSupplier, CellStyle cellStyle, Map<String, String> tdStyle, Map<String, Font> fontMap, CustomColor customColor) {
        Font font = getFont(tdStyle, fontMap, fontSupplier, customColor);
        if (font != null) {
            cellStyle.setFont(font);
        }
    }

    public static Font getFont(Map<String, String> style, Map<String, Font> fontMap, Supplier<Font> fontSupplier, CustomColor customColor) {
        String cacheKey = getCacheKey(style);
        if (fontMap.get(cacheKey) != null) {
            return fontMap.get(cacheKey);
        }
        Font font = null;
        String fs = style.get(FONT_SIZE);
        if (fs != null) {
            short fontSize = (short) TdUtil.getValue(fs);
            font = fontSupplier.get();
            font.setFontHeightInPoints(fontSize);
        }
        String fontFamily = style.get(FONT_FAMILY);
        if (fontFamily != null) {
            font = createFontIfNull(fontSupplier, font);
            font.setFontName(fontFamily);
        }
        String italic = style.get(FONT_STYLE);
        if (ITALIC.equals(italic)) {
            font = createFontIfNull(fontSupplier, font);
            font.setItalic(true);
        }
        String textDecoration = style.get(TEXT_DECORATION);
        if (LINE_THROUGH.equals(textDecoration)) {
            font = createFontIfNull(fontSupplier, font);
            font.setStrikeout(true);
        } else if (UNDERLINE.equals(textDecoration)) {
            font = createFontIfNull(fontSupplier, font);
            font.setUnderline(Font.U_SINGLE);
        }
        String fontWeight = style.get(FONT_WEIGHT);
        if (BOLD.equals(fontWeight)) {
            font = createFontIfNull(fontSupplier, font);
            font.setBold(true);
        }
        String fontColor = style.get(FONT_COLOR);
        if (StringUtil.isNotBlank(fontColor)) {
            font = createFontIfNull(fontSupplier, font);
            font = setFontColor(font, customColor, fontColor);
        }
        if (font != null) {
            fontMap.put(cacheKey, font);
        }
        return font;
    }

    private static Font setFontColor(Font font, CustomColor customColor, String fontColor) {
        Short colorPredefined = ColorUtil.getPredefinedColorIndex(fontColor);
        if (colorPredefined != null) {
            font.setColor(colorPredefined);
            return font;
        }
        int[] rgb = ColorUtil.getRGBByColor(fontColor);
        if (rgb == null) {
            return null;
        }
        if (customColor.isXls()) {
            short index = ColorUtil.getCustomColorIndex(customColor, rgb);
            font.setColor(index);
        } else {
            ((XSSFFont) font).setColor(new XSSFColor(new Color(rgb[0], rgb[1], rgb[2]), customColor.getDefaultIndexedColorMap()));
        }
        return font;
    }

    private static Font createFontIfNull(Supplier<Font> fontSupplier, Font font) {
        return font == null ? fontSupplier.get() : font;
    }

    private static String getCacheKey(Map<String, String> tdStyle) {
        StringBuilder result = new StringBuilder();
        appendKey(tdStyle, FONT_SIZE, result);
        appendKey(tdStyle, FONT_FAMILY, result);
        appendKey(tdStyle, FONT_STYLE, result);
        appendKey(tdStyle, TEXT_DECORATION, result);
        appendKey(tdStyle, FONT_WEIGHT, result);
        appendKey(tdStyle, FONT_COLOR, result);
        return result.toString();
    }

    private static void appendKey(Map<String, String> tdStyle, String styleName, StringBuilder result) {
        String style = tdStyle.get(styleName);
        if (style != null) {
            result.append(styleName).append(":").append(style).append("_");
        }
    }
}
