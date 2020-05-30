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

import java.lang.reflect.Field;

/**
 * 读取异常上下文
 *
 * @author liaochong
 * @version 1.0
 */
public class ReadContext<T> {

    private T object;

    private Field field;

    private String val;

    private int rowNum;

    private int colNum;

    public void reset(T object, Field field, String val, int rowNum, int colNum) {
        this.object = object;
        this.field = field;
        this.val = val;
        this.rowNum = rowNum;
        this.colNum = colNum;
    }

    public T getObject() {
        return this.object;
    }

    public Field getField() {
        return this.field;
    }

    public String getVal() {
        return this.val;
    }

    public int getRowNum() {
        return this.rowNum;
    }

    public int getColNum() {
        return this.colNum;
    }

    public void setObject(T object) {
        this.object = object;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public void setVal(String val) {
        this.val = val;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public void setColNum(int colNum) {
        this.colNum = colNum;
    }
}
