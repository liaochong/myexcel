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
