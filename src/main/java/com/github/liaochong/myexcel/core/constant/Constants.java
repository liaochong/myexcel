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
package com.github.liaochong.myexcel.core.constant;

import com.github.liaochong.myexcel.core.container.Pair;

import javax.lang.model.type.NullType;

/**
 * 常量集合
 *
 * @author liaochong
 * @version 1.0
 */
public class Constants {

    public static final String XLS = ".xls";

    public static final String XLSX = ".xlsx";

    public static final String TRUE = "true";

    public static final String FALSE = "false";

    public static final String ONE = "1";

    public static final String ZERO = "0";

    public static final String HTML_SUFFIX = ".html";

    public static final String HTTP = "http";

    public static final String DATA = "data";

    public static final String COMMA = ",";

    public static final String QUOTES = "\"";

    public static final String CSV = ".csv";

    public static final String COLON = ":";

    public static final String ARROW = "->";

    public static final String SPOT = ".";

    public static final String LEFT_BRACKET = "(";

    public static final String RIGHT_BRACKET = ")";

    public static final String EQUAL = "=";

    public static final Pair<Class, Object> NULL_PAIR = Pair.of(NullType.class, null);
}
