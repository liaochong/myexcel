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
package com.github.liaochong.myexcel.core;

import javax.validation.ConstraintViolation;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
public class ValidationListObject<T> {

    private List<ValidationObject<T>> validationObjects = new LinkedList<>();

    public List<ValidationObject<T>> getValidationObjects() {
        return validationObjects;
    }

    public void setValidationObjects(List<ValidationObject<T>> validationObjects) {
        this.validationObjects = validationObjects;
    }

    public List<T> getObjectList() {
        return validationObjects.stream().map(ValidationObject::getObject).collect(Collectors.toList());
    }

    public List<Set<ConstraintViolation<T>>> getViolations() {
        return validationObjects.stream().map(ValidationObject::getConstraintViolations).collect(Collectors.toList());
    }
}
