package com.github.liaochong.html2excel.utils;

import org.jsoup.nodes.Element;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public final class StyleUtils {

    public static Map<String, String> parseStyle(Element element) {
        String style = element.attr("style");
        if (Objects.isNull(style)) {
            return Collections.emptyMap();
        }
        String[] styleArr = style.split(";");

        Map<String, String> result = new HashMap<>(styleArr.length);
        for (int i = 0, length = styleArr.length; i < length; i++) {
            String[] styleDetail = styleArr[i].split(":");
            if (styleDetail.length < 2) {
                continue;
            }
            String styleName = styleDetail[0].trim();
            if (styleName.length() == 0) {
                continue;
            }
            String styleValue = styleDetail[1].trim();
            if (styleValue.length() == 0) {
                continue;
            }
            result.put(styleName, styleValue);
        }
        return result;
    }
}
