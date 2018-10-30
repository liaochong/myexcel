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
        Font font = workbook.createFont();

        String fs = tdStyle.get("font-size");
        if (Objects.nonNull(fs)) {
            fs = fs.replaceAll("\\D*", "");
            short fontSize = Short.parseShort(fs);
            font.setFontHeightInPoints(fontSize);
        }

        String fontFamily = tdStyle.get("font-family");
        if (Objects.nonNull(fontFamily)) {
            font.setFontName(fontFamily);
        }
        String italic = tdStyle.get("font-style");
        if (Objects.equals("italic", italic)) {
            font.setItalic(true);
        }
        String strikeout = tdStyle.get("text-decoration");
        if (Objects.equals(strikeout, "line-through")) {
            font.setStrikeout(true);
        }
        cellStyle.setFont(font);
    }
}
