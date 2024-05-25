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

/**
 * @author liaochong
 * @version 1.0
 */
public class ValidationObject<T> {

    private List<T> objects = new LinkedList<>();

    private List<ValidationInfo<T>> validationInfos = new LinkedList<>();

    public List<T> getObjects() {
        return objects;
    }

    public void setObjects(List<T> objects) {
        this.objects = objects;
    }

    public List<ValidationInfo<T>> getValidationInfos() {
        return validationInfos;
    }

    public void setValidationInfos(List<ValidationInfo<T>> validationInfos) {
        this.validationInfos = validationInfos;
    }

    public static class ValidationInfo<T> {
        private int rowNum;

        private Set<ConstraintViolation<T>> constraintViolations;

        public int getRowNum() {
            return rowNum;
        }

        public void setRowNum(int rowNum) {
            this.rowNum = rowNum;
        }

        public Set<ConstraintViolation<T>> getConstraintViolations() {
            return constraintViolations;
        }

        public void setConstraintViolations(Set<ConstraintViolation<T>> constraintViolations) {
            this.constraintViolations = constraintViolations;
        }
    }
}
