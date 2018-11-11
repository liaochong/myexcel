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

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class LRU<K, V> {

    private static final float LOAD_FACTOR = 0.75f;
    /**
     * 默认存储数量
     */
    private static final int DEFAULT_SIZE = 2 << 8;
    /**
     * 最大存储数量
     */
    private int maxCapacity;

    private LinkedHashMap<K, V> map;

    public LRU() {
        this.map = new LinkedHashMap<K, V>(DEFAULT_SIZE, LOAD_FACTOR, true) {
            @Override
            protected boolean removeEldestEntry(Map.Entry eldest) {
                return LRU.this.map.size() == maxCapacity;
            }
        };
    }

    public LRU(int maxCapacity) {
        if (maxCapacity < 1) {
            throw new IllegalArgumentException("Capacity must be greater than 0");
        }
        this.maxCapacity = maxCapacity;
        this.map = new LinkedHashMap<K, V>(maxCapacity, LOAD_FACTOR, true) {
            @Override
            protected boolean removeEldestEntry(Map.Entry eldest) {
                return LRU.this.map.size() == maxCapacity;
            }
        };
    }

    public synchronized void put(K key, V val) {
        this.map.put(key, val);
    }

    public V get(K key) {
        return this.map.get(key);
    }

    public void remove(K key) {
        this.map.remove(key);
    }

    public synchronized void clear() {
        this.map.clear();
    }
}
