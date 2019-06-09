/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.style;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Map;
import java.util.Objects;

/**
 * 换行样式
 *
 * @author liaochong
 * @version 1.0
 */
public class WordBreakStyle {

    public static final String WORD_BREAK = "word-break";

    public static final String BREAK_ALL = "break-all";

    public static void setWordBreak(CellStyle cellStyle, Map<String, String> tdStyle) {
        String wordBreak = tdStyle.get(WORD_BREAK);
        if (Objects.equals(BREAK_ALL, wordBreak)) {
            cellStyle.setWrapText(true);
        }
    }
}
