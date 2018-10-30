package com.github.liaochong.html2excel.core.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class TextAlignStyle {

    private static Map<String, HorizontalAlignment> horizontalAlignmentMap;

    private static Map<String, VerticalAlignment> verticalAlignmentMap;

    static {
        horizontalAlignmentMap = Arrays.stream(HorizontalAlignment.values()).collect(Collectors.toMap(h -> h.name().toLowerCase(), h -> h));
        verticalAlignmentMap = Arrays.stream(VerticalAlignment.values()).collect(Collectors.toMap(v -> v.name().toLowerCase(), v -> v));
        verticalAlignmentMap.put("middle", VerticalAlignment.CENTER);
    }

    public static void setTextAlign(CellStyle cellStyle, Map<String, String> tdStyle) {
        if (Objects.isNull(tdStyle)) {
            return;
        }
        String textAlign = tdStyle.get("text-align");
        if (horizontalAlignmentMap.containsKey(textAlign)) {
            cellStyle.setAlignment(horizontalAlignmentMap.get(textAlign));
        }
        String verticalAlign = tdStyle.get("vertical-align");
        if (verticalAlignmentMap.containsKey(verticalAlign)) {
            cellStyle.setVerticalAlignment(verticalAlignmentMap.get(verticalAlign));
        }
    }
}
