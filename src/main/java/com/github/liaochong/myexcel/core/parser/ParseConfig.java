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

import com.github.liaochong.myexcel.core.strategy.AutoWidthStrategy;
import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

/**
 * 解析配置
 *
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ParseConfig {

    AutoWidthStrategy autoWidthStrategy;

    boolean isCustomWidth;

    boolean isComputeAutoWidth;

    public void setAutoWidthStrategy(AutoWidthStrategy autoWidthStrategy) {
        this.autoWidthStrategy = autoWidthStrategy;
        this.isCustomWidth = AutoWidthStrategy.isCustomWidth(autoWidthStrategy);
        this.isComputeAutoWidth = AutoWidthStrategy.isComputeAutoWidth(autoWidthStrategy);
    }
}
