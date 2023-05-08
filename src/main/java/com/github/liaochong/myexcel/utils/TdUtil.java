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
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.cache.Cache;
import com.github.liaochong.myexcel.core.cache.WeakCache;

import java.util.regex.Pattern;


/**
 * @author liaochong
 * @version 1.0
 */
public final class TdUtil {

    private static final Pattern chineseOrCapitalPattern = Pattern.compile("[\u4e00-\u9fa5|A-Z]");

    private static final Pattern digitalPattern = Pattern.compile("^\\d+$");

    private static final Pattern nonDigitalPattern = Pattern.compile("[^\\d]+");

    private static final Cache<String, Integer> SPAN_CACHE = new WeakCache<>();

    public static int get(int firstValue, int secondValue) {
        return firstValue > 0 ? secondValue + firstValue - 1 : secondValue;
    }

    public static int getSpan(String span) {
        Integer cacheResult = SPAN_CACHE.get(span);
        if (cacheResult != null) {
            return cacheResult;
        }
        if (!isSpanValid(span)) {
            SPAN_CACHE.cache(span, 0);
            return 0;
        }
        int result = Integer.parseInt(span);
        SPAN_CACHE.cache(span, result);
        return result;
    }

    public static boolean isSpanValid(String span) {
        return digitalPattern.matcher(span).find();
    }

    public static int getStringWidth(String s) {
        return getStringWidth(s, 0);
    }

    public static int getStringWidth(String s, double shift) {
        if (s == null) {
            return 1;
        }
        // 最小为1
        double valueLength = 1;
        // 获取字段值的长度，如果含中文字符，则每个中文字符长度为1，否则为0.5
        double chineseOrCapitalShift = 1 + shift;
        double otherShift = 0.5 + shift;
        for (int i = 0, size = s.length(); i < size; i++) {
            // 获取一个字符
            String temp = s.substring(i, i + 1);
            if (chineseOrCapitalPattern.matcher(temp).find()) {
                valueLength += chineseOrCapitalShift;
            } else {
                valueLength += otherShift;
            }
        }
        // 进位取整
        return (int) Math.ceil(valueLength);
    }

    public static int getValue(String v) {
        if (v == null) {
            return -1;
        }
        String realValue = nonDigitalPattern.matcher(v).replaceAll("");
        return Integer.parseInt(realValue);
    }
}
