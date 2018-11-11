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
package com.github.liaochong.html2excel.core.cache;

import org.apache.poi.ss.usermodel.Font;

import java.util.Objects;

/**
 * 字体缓存
 *
 * @author liaochong
 * @version 1.0
 */
public class FontCache {

    private static LRU<String, Font> lru = new LRU<>();

    public static void cache(String key, Font value) {
        lru.put(key, value);
    }

    public static Font get(String key) {
        return lru.get(key);
    }

    public static boolean contains(String key) {
        return Objects.nonNull(lru.get(key));
    }
}
