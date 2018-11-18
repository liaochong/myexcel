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

import com.github.liaochong.html2excel.core.cache.DefaultCache;
import org.jsoup.nodes.Element;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * 样式工具
 *
 * @author liaochong
 * @version 1.0
 */
public final class StyleUtils {

    private static final DefaultCache<String, Map<String, String>> STYLE_CACHE = new DefaultCache<>();

    public static Map<String, String> parseStyle(Element element) {
        String style = element.attr("style");
        Map<String, String> cacheResult = STYLE_CACHE.get(style);
        if (Objects.nonNull(cacheResult)) {
            return cacheResult;
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
        STYLE_CACHE.cache(style, result);
        return result;
    }

    /**
     * 样式融合
     *
     * @param originStyle 源样式
     * @param targetStyle 目标样式
     * @return 结果
     */
    public static Map<String, String> mixStyle(Map<String, String> originStyle, Map<String, String> targetStyle) {
        if (Objects.isNull(targetStyle) && Objects.isNull(originStyle)) {
            return Collections.emptyMap();
        }
        if (Objects.equals(originStyle, targetStyle)) {
            return targetStyle;
        }
        if (Objects.isNull(targetStyle)) {
            return originStyle;
        } else {
            originStyle.forEach(targetStyle::putIfAbsent);
            return targetStyle;
        }
    }
}
