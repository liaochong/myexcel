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

import java.util.Arrays;
import java.util.Objects;

/**
 * 线形枚举
 *
 * @author liaochong
 * @version 1.0
 */
public enum LineStyleEnum {

    SOLID(0),

    DOT(1),

    DASH(2),

    LG_DASH(3),

    DASH_DOT(4),

    LG_DASH_DOT(5),

    LG_DASH_DOT_DOT(6),

    SYS_DASH(7),

    SYS_DOT(8),

    SYS_DASH_DOT(9),

    SYS_DASH_DOT_DOT(10);

    private final int value;

    LineStyleEnum(int value) {
        this.value = value;
    }

    public int getValue() {
        return value;
    }

    public static int getByName(String name) {
        String upperCaseName = name.toUpperCase();
        return Arrays.stream(LineStyleEnum.values()).filter(lineStyleEnum -> Objects.equals(lineStyleEnum.name(), upperCaseName)).findFirst()
                .orElseThrow(() -> new IllegalArgumentException("No matching line type found."))
                .value;
    }
}
