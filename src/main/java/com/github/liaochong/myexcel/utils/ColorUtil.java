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
package com.github.liaochong.myexcel.utils;

import com.github.liaochong.myexcel.core.CustomColor;
import lombok.NonNull;
import lombok.experimental.UtilityClass;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.util.HSSFColor;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * @author liaochong
 * @version 1.0
 */
@UtilityClass
public class ColorUtil {

    private static final String HASH = "#";

    private static final String RGB = "rgb";

    private static final Map<String, HSSFColor.HSSFColorPredefined> COLOR_PREDEFINED_MAP;

    static {
        COLOR_PREDEFINED_MAP = Arrays.stream(HSSFColor.HSSFColorPredefined.values())
                .collect(Collectors.toMap(c -> c.toString().toLowerCase().replaceAll("_", ""), c -> c));
    }

    public static Short getPredefinedColorIndex(String color) {
        HSSFColor.HSSFColorPredefined colorPredefined = COLOR_PREDEFINED_MAP.get(color);
        if (Objects.isNull(colorPredefined)) {
            return null;
        }
        return colorPredefined.getIndex();
    }

    public static Short getCustomColorIndex(@NonNull CustomColor customColor, @NonNull String color) {
        int[] rgb = getRGBByColor(color);
        if (Objects.isNull(rgb)) {
            return null;
        }
        return getCustomColorIndex(customColor, rgb);
    }

    public static Short getCustomColorIndex(@NonNull CustomColor customColor, @NonNull int[] rgb) {
        HSSFPalette palette = customColor.getPalette();
        short index = (short) customColor.getColorIndex().getAndIncrement();
        palette.setColorAtIndex(index, (byte) rgb[0], (byte) rgb[1], (byte) rgb[2]);
        return index;
    }

    public static int[] getRGBByColor(@NonNull String color) {
        int[] result = null;
        if (color.startsWith(HASH)) {
            // 转为16进制
            int r = Integer.parseInt((color.substring(1, 3)), 16);
            int g = Integer.parseInt((color.substring(3, 5)), 16);
            int b = Integer.parseInt((color.substring(5, 7)), 16);

            result = new int[3];
            result[0] = r;
            result[1] = g;
            result[2] = b;
        } else if (color.startsWith(RGB)) {
            String rgbColor = color.replace(RGB, "").replace("(", "").replace(")", "");
            String[] rgbColorArr = rgbColor.split(",");
            List<Integer> rgb = Arrays.stream(rgbColorArr)
                    .map(String::trim).map(Integer::parseInt)
                    .collect(Collectors.toList());
            if (rgb.size() != 3) {
                return result;
            }
            // 转为16进制
            int r = rgb.get(0);
            int g = rgb.get(1);
            int b = rgb.get(2);

            result = new int[3];
            result[0] = r;
            result[1] = g;
            result[2] = b;
        }
        return result;
    }

}
