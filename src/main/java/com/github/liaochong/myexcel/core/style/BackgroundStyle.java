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
package com.github.liaochong.myexcel.core.style;

import com.github.liaochong.myexcel.core.CustomColor;
import com.github.liaochong.myexcel.utils.ColorUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.awt.*;
import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
public final class BackgroundStyle {

    public static final String BACKGROUND_COLOR = "background-color";


    public static void setBackgroundColor(CellStyle style, Map<String, String> tdStyle, CustomColor customColor) {
        if (Objects.isNull(tdStyle)) {
            return;
        }
        String color = tdStyle.get(BACKGROUND_COLOR);
        if (Objects.isNull(color)) {
            return;
        }
        Short colorPredefined = ColorUtil.getPredefinedColorIndex(color);
        if (Objects.nonNull(colorPredefined)) {
            style.setFillForegroundColor(colorPredefined);
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return;
        }
        byte[] rgb = ColorUtil.getRGBByColor(color);
        setCustomColor(style, rgb, customColor);
    }

    private static void setCustomColor(CellStyle style, byte[] rgb, CustomColor customColor) {
        if (Objects.isNull(rgb)) {
            return;
        }
        if (customColor.isXls()) {
            short index = ColorUtil.getCustomColorIndex(customColor, rgb);
            style.setFillForegroundColor(index);
        } else {
            XSSFCellStyle xssfCellStyle = (XSSFCellStyle) style;
            xssfCellStyle.setFillForegroundColor(new XSSFColor(new Color(rgb[0], rgb[1], rgb[2]), customColor.getDefaultIndexedColorMap()));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
    }

}
