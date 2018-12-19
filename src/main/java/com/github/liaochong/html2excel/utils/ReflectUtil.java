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

import com.github.liaochong.html2excel.core.reflect.ClassFieldContainer;

import java.lang.reflect.Field;

/**
 * @author liaochong
 * @version 1.0
 */
public final class ReflectUtil {

    public static ClassFieldContainer getAllFieldsOfClass(Class<?> clazz) {
        ClassFieldContainer container = new ClassFieldContainer();
        getAllFieldsOfClass(clazz, container);
        return container;
    }

    private static void getAllFieldsOfClass(Class<?> clazz, ClassFieldContainer container) {
        container.setClazz(clazz);
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
            container.getFields().add(field);
            container.getFieldMap().put(field.getName(), field);
        }
        if (clazz.getSuperclass() != null) {
            ClassFieldContainer parentContainer = new ClassFieldContainer();
            container.setParent(parentContainer);
            getAllFieldsOfClass(clazz.getSuperclass(), parentContainer);
        }
    }
}
