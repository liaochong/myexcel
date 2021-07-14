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

/**
 * 斜线
 *
 * @author liaochong
 * @version 1.0
 */
public class Slant {
    /**
     * 线宽
     */
    private double lineWidth = 0.5;
    /**
     * 线的风格
     */
    private int lineStyle = 0;
    /**
     * 线的颜色
     */
    private String lineStyleColor = "#000000";

    public Slant() {
    }

    public Slant(int lineStyle, String lineWidth, String lineStyleColor) {
        this.lineWidth = Double.parseDouble(lineWidth);
        this.lineStyle = lineStyle;
        this.lineStyleColor = lineStyleColor;
    }

    public double getLineWidth() {
        return lineWidth;
    }

    public void setLineWidth(double lineWidth) {
        this.lineWidth = lineWidth;
    }

    public int getLineStyle() {
        return lineStyle;
    }

    public void setLineStyle(int lineStyle) {
        this.lineStyle = lineStyle;
    }

    public String getLineStyleColor() {
        return lineStyleColor;
    }

    public void setLineStyleColor(String lineStyleColor) {
        this.lineStyleColor = lineStyleColor;
    }
}
