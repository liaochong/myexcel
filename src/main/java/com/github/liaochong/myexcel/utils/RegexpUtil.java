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
package com.github.liaochong.myexcel.utils;

import java.util.regex.Pattern;

/**
 * @author liaochong
 * @version 1.0
 */
public final class RegexpUtil {

    private static final Pattern PATTERN_COMMA = Pattern.compile(",");

    private static final Pattern LINE_FEED_PATTERN = Pattern.compile("\n");

    private static final Pattern UN_LINE_FEED_PATTERN = Pattern.compile("&#10;");

    public static String removeComma(String content) {
        if (content == null) {
            return content;
        }
        return PATTERN_COMMA.matcher(content).replaceAll("");
    }

    public static String escapeLineFeed(String content) {
        if (content == null) {
            return null;
        }
        return LINE_FEED_PATTERN.matcher(content).replaceAll("&#10;");
    }

    public static String unescapeLineFeed(String content) {
        if (content == null) {
            return null;
        }
        return UN_LINE_FEED_PATTERN.matcher(content).replaceAll("\n");
    }
}
