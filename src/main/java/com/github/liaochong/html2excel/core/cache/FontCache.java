package com.github.liaochong.html2excel.core.cache;

import org.apache.poi.ss.usermodel.Font;

import java.util.Objects;

/**
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
