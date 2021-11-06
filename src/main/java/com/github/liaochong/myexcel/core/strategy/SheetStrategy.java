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
package com.github.liaochong.myexcel.core.strategy;

import java.util.Objects;

/**
 * table 策略
 *
 * @author QingMings
 * @since v3.11.3
 */
public enum SheetStrategy {
    /**
     * 多个table生成在同一个sheet里
     */
    ONE_SHEET,
    /**
     * 每个table各生成一个sheet
     */
    MULTI_SHEET;

    public static boolean isOneSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, ONE_SHEET);
    }

    public static boolean isMultiSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, MULTI_SHEET);
    }
}
