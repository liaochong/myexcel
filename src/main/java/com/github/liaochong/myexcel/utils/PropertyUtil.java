/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.ExcelColumnMapping;
import com.github.liaochong.myexcel.core.cache.WeakCache;
import com.github.liaochong.myexcel.core.constant.Constants;

import java.util.Properties;

/**
 * 配置工具
 *
 * @author liaochong
 * @version 1.0
 */
public final class PropertyUtil {

    private static WeakCache<ExcelColumnMapping, Properties> mappingCache = new WeakCache<>();

    private static WeakCache<ExcelColumnMapping, Properties> reverseMappingCache = new WeakCache<>();

    private static final Properties EMPTY_PROPERTIES = new Properties();

    public static java.util.Properties getProperties(ExcelColumnMapping excelColumnMapping) {
        return getProperties(excelColumnMapping, mappingCache, false);
    }

    public static java.util.Properties getReverseProperties(ExcelColumnMapping excelColumnMapping) {
        return getProperties(excelColumnMapping, reverseMappingCache, true);
    }

    private static java.util.Properties getProperties(ExcelColumnMapping excelColumnMapping, WeakCache<ExcelColumnMapping, Properties> mappingCache, boolean reverse) {
        Properties properties = mappingCache.get(excelColumnMapping);
        if (properties != null) {
            return properties;
        }
        String[] mappingGroups = excelColumnMapping.getMapping().split(Constants.COMMA);
        if (mappingGroups.length == 0) {
            mappingCache.cache(excelColumnMapping, EMPTY_PROPERTIES);
            return EMPTY_PROPERTIES;
        }
        properties = new java.util.Properties();
        for (String m : mappingGroups) {
            String[] mappingGroup = m.split(Constants.COLON);
            if (mappingGroup.length != 2) {
                throw new IllegalArgumentException("Illegal mapping:" + m);
            }
            if (reverse) {
                properties.setProperty(mappingGroup[1], mappingGroup[0]);
            } else {
                properties.setProperty(mappingGroup[0], mappingGroup[1]);
            }
        }
        mappingCache.cache(excelColumnMapping, properties);
        return properties;
    }
}
