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

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.Arrays;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public final class TextAlignStyle {

    public static final String TEXT_ALIGN = "text-align";

    public static final String VERTICAL_ALIGN = "vertical-align";

    public static final String MIDDLE = "middle";

    public static final String CENTER = "center";

    private static Map<String, HorizontalAlignment> horizontalAlignmentMap;

    private static Map<String, VerticalAlignment> verticalAlignmentMap;

    static {
        horizontalAlignmentMap = Arrays.stream(HorizontalAlignment.values()).collect(Collectors.toMap(h -> h.name().toLowerCase(), h -> h));
        verticalAlignmentMap = Arrays.stream(VerticalAlignment.values()).collect(Collectors.toMap(v -> v.name().toLowerCase(), v -> v));
        verticalAlignmentMap.put(MIDDLE, VerticalAlignment.CENTER);
    }

    public static void setTextAlign(CellStyle cellStyle, Map<String, String> tdStyle) {
        if (tdStyle == null) {
            return;
        }
        String textAlign = tdStyle.get(TEXT_ALIGN);
        if (horizontalAlignmentMap.containsKey(textAlign)) {
            cellStyle.setAlignment(horizontalAlignmentMap.get(textAlign));
        }
        String verticalAlign = tdStyle.get(VERTICAL_ALIGN);
        if (verticalAlignmentMap.containsKey(verticalAlign)) {
            cellStyle.setVerticalAlignment(verticalAlignmentMap.get(verticalAlign));
        }
    }
}
