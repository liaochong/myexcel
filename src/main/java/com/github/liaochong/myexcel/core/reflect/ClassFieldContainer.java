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
package com.github.liaochong.myexcel.core.reflect;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class ClassFieldContainer {

    private Class<?> clazz;

    private final List<Field> declaredFields = new ArrayList<>();

    private final Map<String, Field> fieldMap = new HashMap<>();

    private ClassFieldContainer parent;

    public Field getFieldByName(String fieldName) {
        return this.getFieldByName(fieldName, this);
    }

    @SafeVarargs
    public final List<Field> getFieldsByAnnotation(Class<? extends Annotation>... annotationClazzs) {
        Objects.requireNonNull(annotationClazzs);
        List<Field> annotationFields = new ArrayList<>();
        this.getFieldsByAnnotation(this, annotationFields, annotationClazzs);
        return annotationFields;
    }

    public List<Field> getFields() {
        List<Field> fields = new ArrayList<>();
        this.getFieldsByContainer(this, fields);
        return fields;
    }

    private void getFieldsByContainer(ClassFieldContainer classFieldContainer, List<Field> fields) {
        ClassFieldContainer parentContainer = classFieldContainer.getParent();
        if (parentContainer != null) {
            this.getFieldsByContainer(parentContainer, fields);
        }
        filterFields(classFieldContainer.getDeclaredFields(), fields);
    }

    @SafeVarargs
    private final void getFieldsByAnnotation(ClassFieldContainer classFieldContainer, List<Field> annotationFieldContainer, Class<? extends Annotation>... annotationClazzs) {
        ClassFieldContainer parentContainer = classFieldContainer.getParent();
        if (parentContainer != null) {
            this.getFieldsByAnnotation(parentContainer, annotationFieldContainer, annotationClazzs);
        }
        List<Field> annotationFields = classFieldContainer.declaredFields.stream().filter(field -> {
            for (Class<? extends Annotation> annotationClazz : annotationClazzs) {
                boolean isAnnotationPresent = field.isAnnotationPresent(annotationClazz);
                if (isAnnotationPresent) {
                    return true;
                }
            }
            return false;
        }).collect(Collectors.toList());
        filterFields(annotationFields, annotationFieldContainer);
    }

    private void filterFields(List<Field> declaredFields, List<Field> fieldContainer) {
        to:
        for (Field field : declaredFields) {
            for (int j = 0; j < fieldContainer.size(); j++) {
                Field f = fieldContainer.get(j);
                if (f.getName().equals(field.getName())) {
                    fieldContainer.set(j, field);
                    continue to;
                }
            }
            fieldContainer.add(field);
        }
    }

    private Field getFieldByName(String fieldName, ClassFieldContainer container) {
        Field field = container.getFieldMap().get(fieldName);
        if (field != null) {
            return field;
        }
        ClassFieldContainer parentContainer = container.getParent();
        if (parentContainer == null) {
            return null;
        }
        return getFieldByName(fieldName, parentContainer);
    }

    public Class<?> getClazz() {
        return this.clazz;
    }

    public List<Field> getDeclaredFields() {
        return this.declaredFields;
    }

    public Map<String, Field> getFieldMap() {
        return this.fieldMap;
    }

    public ClassFieldContainer getParent() {
        return this.parent;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public void setParent(ClassFieldContainer parent) {
        this.parent = parent;
    }
}
