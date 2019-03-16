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

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;

import java.util.Objects;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class CustomColor {

    boolean isXls = false;

    HSSFPalette palette;

    AtomicInteger colorIndex = new AtomicInteger(56);

    DefaultIndexedColorMap defaultIndexedColorMap;

    public CustomColor(boolean isXls, HSSFPalette palette) {
        this.isXls = isXls;
        this.palette = palette;
    }

    public CustomColor() {
    }

    public DefaultIndexedColorMap getDefaultIndexedColorMap() {
        if (Objects.isNull(defaultIndexedColorMap)) {
            defaultIndexedColorMap = new DefaultIndexedColorMap();
        }
        return defaultIndexedColorMap;
    }
}
