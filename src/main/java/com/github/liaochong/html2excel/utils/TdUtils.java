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
package com.github.liaochong.html2excel.utils;

import java.util.Objects;
import java.util.function.IntSupplier;
import java.util.regex.Pattern;


/**
 * @author liaochong
 * @version 1.0
 */
public class TdUtils {

    private static Pattern pattern = Pattern.compile("^\\d+$");

    public static int get(IntSupplier firstSupplier, IntSupplier secondSupplier) {
        int firstValue = firstSupplier.getAsInt();
        int secondValue = secondSupplier.getAsInt();
        return firstValue > 0 ? secondValue + firstValue - 1 : secondValue;
    }

    public static int getSpan(String span) {
        if (!isSpanValid(span)) {
            return 0;
        }
        int result = Integer.parseInt(span);
        return result > 1 ? result : 0;
    }

    public static boolean isSpanValid(String span) {
        return pattern.matcher(span).find();
    }

    public static int getStringWidth(String s) {
        if (Objects.isNull(s)) {
            return 1;
        }
        // 最小为1
        double valueLength = 1;
        String chinese = "[\u4e00-\u9fa5]";
        // 获取字段值的长度，如果含中文字符，则每个中文字符长度为2，否则为1
        for (int i = 0; i < s.length(); i++) {
            // 获取一个字符
            String temp = s.substring(i, i + 1);
            // 判断是否为中文字符
            if (temp.matches(chinese)) {
                // 中文字符长度为1
                valueLength += 1;
            } else {
                // 其他字符长度为0.5
                valueLength += 0.5;
            }
        }
        // 进位取整
        return (int) Math.ceil(valueLength);
    }

}
