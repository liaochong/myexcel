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
package com.github.liaochong.myexcel.core.cache;

import java.util.WeakHashMap;

/**
 * @author liaochong
 * @version 1.0
 */
public class WeakCache<K, V> implements Cache<K, V> {

    private volatile WeakHashMap<K, V> cacheMap;

    public WeakCache() {
        cacheMap = new WeakHashMap<>();
    }

    public WeakCache(int mapSize) {
        cacheMap = new WeakHashMap<>(mapSize);
    }

    @Override
    public synchronized void cache(K key, V value) {
        cacheMap.put(key, value);
    }

    @Override
    public V get(K key) {
        return cacheMap.get(key);
    }

    @Override
    public synchronized void clearAll() {
        cacheMap.clear();
    }
}
