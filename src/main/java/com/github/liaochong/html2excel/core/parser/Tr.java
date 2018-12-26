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
package com.github.liaochong.html2excel.core.parser;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;
import org.jsoup.nodes.Element;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class Tr {

    Element element;
    /**
     * 索引
     */
    int index;
    /**
     * 行单元格
     */
    List<Td> tdList = new ArrayList<>();
    /**
     * 行样式
     */
    Map<String, String> style;
    /**
     * 当前最后列编号
     */
    int lastColumnNum;
    /**
     * 最大宽度
     */
    Map<Integer, Integer> colWidthMap;

    public Tr(int index) {
        this.index = index;
    }
}
