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
public class DefaultCache<E, T> implements Cache<E, T> {

    private static final int DEFAULT_CACHE_SIZE = 500;

    private int cacheSize = DEFAULT_CACHE_SIZE;

    private LinkedHashMap<E, T> cacheMap = new LinkedHashMap<E, T>() {
        @Override
        protected boolean removeEldestEntry(Map.Entry<E, T> eldest) {
            return this.size() > cacheSize;
        }
    };

    public DefaultCache() {
    }

    public DefaultCache(int cacheSize) {
        this.cacheSize = cacheSize;
    }

    @Override
    public synchronized void cache(E key, T value) {
        cacheMap.put(key, value);
    }

    @Override
    public T get(E key) {
        return cacheMap.get(key);
    }

    @Override
    public synchronized void clearAll() {
        cacheMap.clear();
    }
}
