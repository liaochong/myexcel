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
import lombok.experimental.UtilityClass;
import org.jsoup.nodes.Element;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * 样式工具
 *
 * @author liaochong
 * @version 1.0
 */
@UtilityClass
public final class StyleUtil {

    private static final Cache<String, Map<String, String>> STYLE_CACHE = new WeakCache<>();

    public static Map<String, String> parseStyle(Element element) {
        String style = element.attr("style");
        return parseStyle(style);
    }


    public static Map<String, String> parseStyle(String style) {
        if (style.length() == 0) {
            return Collections.emptyMap();
        }
        Map<String, String> cacheResult = STYLE_CACHE.get(style);
        if (cacheResult != null) {
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
        if (targetStyle == null && originStyle == null) {
            return Collections.emptyMap();
        }
        if (targetStyle == null) {
            return new HashMap<>(originStyle);
        } else if (originStyle == null) {
            return new HashMap<>(targetStyle);
        }
        // 相加的两倍，防止扩容。
        Map<String, String> result = new HashMap<>((targetStyle.size() + originStyle.size()) * 2);
        originStyle.forEach(result::putIfAbsent);
        targetStyle.forEach(result::put);
        return result;
    }
}
