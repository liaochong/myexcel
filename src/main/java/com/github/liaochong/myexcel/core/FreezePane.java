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
package com.github.liaochong.myexcel.core;

/**
 * @author liaochong
 * @version 1.0
 */
public class FreezePane {

    /**
     * 从左到右需固定列数
     */
    final int colSplit;

    /**
     * 从上到下需固定行数
     */
    final int rowSplit;

    public FreezePane(int rowSplit, int colSplit) {
        this.colSplit = colSplit;
        this.rowSplit = rowSplit;
    }
}
