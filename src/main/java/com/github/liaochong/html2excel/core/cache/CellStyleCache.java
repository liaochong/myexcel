package com.github.liaochong.html2excel.core.cache;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public class CellStyleCache {

    private static LRU<Map<String, String>, CellStyle> lru = new LRU<>();

    public static void cache(Map<String, String> key, CellStyle value) {
        lru.put(key, value);
    }

    public static CellStyle get(Map<String, String> key) {
        return lru.get(key);
    }


    public static boolean contains(Map<String, String> key) {
        return Objects.nonNull(lru.get(key));
    }

}
