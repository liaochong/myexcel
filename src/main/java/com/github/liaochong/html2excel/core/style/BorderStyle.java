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

import java.util.Arrays;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public final class BorderStyle {

    public static final String BORDER_LEFT_STYLE = "border-left-style";

    public static final String BORDER_RIGHT_STYLE = "border-right-style";

    public static final String BORDER_TOP_STYLE = "border-top-style";

    public static final String BORDER_BOTTOM_STYLE = "border-bottom-style";

    public static final String THIN = "thin";

    private static Map<String, org.apache.poi.ss.usermodel.BorderStyle> borderStyleMap;

    static {
        borderStyleMap = Arrays.stream(org.apache.poi.ss.usermodel.BorderStyle.values())
                .collect(Collectors.toMap(b -> b.toString().toLowerCase(), b -> b));
    }

    public static void setBorder(CellStyle cellStyle, Map<String, String> tdStyle) {
        if (Objects.isNull(tdStyle)) {
            return;
        }
        String borderLeftStyle = tdStyle.get(BORDER_LEFT_STYLE);
        if (borderStyleMap.containsKey(borderLeftStyle)) {
            cellStyle.setBorderLeft(borderStyleMap.get(borderLeftStyle));
        }
        String borderRightStyle = tdStyle.get(BORDER_RIGHT_STYLE);
        if (borderStyleMap.containsKey(borderRightStyle)) {
            cellStyle.setBorderRight(borderStyleMap.get(borderRightStyle));
        }
        String borderTopStyle = tdStyle.get(BORDER_TOP_STYLE);
        if (borderStyleMap.containsKey(borderTopStyle)) {
            cellStyle.setBorderTop(borderStyleMap.get(borderTopStyle));
        }
        String borderBottomStyle = tdStyle.get(BORDER_BOTTOM_STYLE);
        if (borderStyleMap.containsKey(borderBottomStyle)) {
            cellStyle.setBorderBottom(borderStyleMap.get(borderBottomStyle));
        }
    }


}
