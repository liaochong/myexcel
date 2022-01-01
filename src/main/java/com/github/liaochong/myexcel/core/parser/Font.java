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
package com.github.liaochong.myexcel.core.parser;

import java.util.HashMap;
import java.util.Map;

/**
 * 富文本字体
 *
 * @author liaochong
 * @version 1.0
 */
public class Font {

    /**
     * 整体字符串中起始位置
     */
    public int startIndex;

    /**
     * 整体字符串中终点位置，不包含
     */
    public int endIndex;

    /**
     * 单元格样式
     */
    public Map<String, String> style = new HashMap<>();
}
