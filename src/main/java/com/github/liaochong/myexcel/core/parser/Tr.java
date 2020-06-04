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
package com.github.liaochong.myexcel.core.parser;

import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class Tr {

    /**
     * 索引
     */
    private int index;
    /**
     * 行单元格
     */
    private List<Td> tdList = Collections.emptyList();
    /**
     * 最大宽度
     */
    private Map<Integer, Integer> colWidthMap;
    /**
     * 是否可见
     */
    private boolean visibility = true;
    /**
     * 行高度
     */
    private int height;
    /**
     * 是否来源于模板
     */
    private boolean fromTemplate;

    public Tr(int index, int height) {
        this.index = index;
        this.height = height;
    }

    public Tr(int index, int height, boolean fromTemplate) {
        this.index = index;
        this.height = height;
        this.fromTemplate = fromTemplate;
    }

    public int getIndex() {
        return this.index;
    }

    public List<Td> getTdList() {
        return this.tdList;
    }

    public Map<Integer, Integer> getColWidthMap() {
        return this.colWidthMap;
    }

    public boolean isVisibility() {
        return this.visibility;
    }

    public int getHeight() {
        return this.height;
    }

    public boolean isFromTemplate() {
        return this.fromTemplate;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public void setTdList(List<Td> tdList) {
        this.tdList = tdList;
    }

    public void setColWidthMap(Map<Integer, Integer> colWidthMap) {
        this.colWidthMap = colWidthMap;
    }

    public void setVisibility(boolean visibility) {
        this.visibility = visibility;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public void setFromTemplate(boolean fromTemplate) {
        this.fromTemplate = fromTemplate;
    }
}
