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

    public static void setFont(Workbook workbook, CellStyle cellStyle, Map<String, String> tdStyle) {
        Font font = null;

        String fs = tdStyle.get("font-size");
        if (Objects.nonNull(fs)) {
            fs = fs.replaceAll("\\D*", "");
            short fontSize = Short.parseShort(fs);
            font = workbook.createFont();
            font.setFontHeightInPoints(fontSize);
        }

        String fontFamily = tdStyle.get("font-family");
        if (Objects.nonNull(fontFamily)) {
            if (Objects.isNull(font)) {
                font = workbook.createFont();
            }
            font.setFontName(fontFamily);
        }
        String italic = tdStyle.get("font-style");
        if (Objects.equals("italic", italic)) {
            if (Objects.isNull(font)) {
                font = workbook.createFont();
            }
            font.setItalic(true);
        }
        String strikeout = tdStyle.get("text-decoration");
        if (Objects.equals(strikeout, "line-through")) {
            if (Objects.isNull(font)) {
                font = workbook.createFont();
            }
            font.setStrikeout(true);
        }
        if (Objects.nonNull(font)) {
            cellStyle.setFont(font);
        }
    }
}
