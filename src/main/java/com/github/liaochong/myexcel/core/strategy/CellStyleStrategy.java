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
 * @author liaochong
 * @version 1.0
 */
public enum CellStyleStrategy {

    /**
     * 无样式
     */
    NO_STYLE,

    /**
     * 默认样式
     */
    DEFAULT_STYLE,

    /**
     * 自适应宽度、高度
     */
    AUTO_WIDTH_HEIGHT,

    /**
     * 组件调整宽度、高度
     */
    DEFAULT_WIDTH_HEIGHT,

    /**
     * 自定义样式
     */
    CUSTOM_STYLE;

    public static boolean isNoStyle(CellStyleStrategy cellStyleStrategy) {
        return Objects.equals(cellStyleStrategy, NO_STYLE);
    }

    public static boolean isDefaultStyle(CellStyleStrategy cellStyleStrategy) {
        return Objects.equals(cellStyleStrategy, DEFAULT_STYLE);
    }

    public static boolean isCustomStyle(CellStyleStrategy cellStyleStrategy) {
        return Objects.equals(cellStyleStrategy, CUSTOM_STYLE);
    }

    public static boolean isAutoWidthHeight(CellStyleStrategy cellStyleStrategy) {
        return Objects.equals(cellStyleStrategy, AUTO_WIDTH_HEIGHT);
    }

    public static boolean isDefaultWidthHeight(CellStyleStrategy cellStyleStrategy) {
        return Objects.equals(cellStyleStrategy, DEFAULT_WIDTH_HEIGHT);
    }
}
