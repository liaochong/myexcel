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
package com.github.liaochong.myexcel.core.watermark;

import java.awt.*;

/**
 * 水印
 *
 * @author liaochong
 * @version 1.0
 */
public class Watermark {
    /**
     * 水印文本
     */
    private String text;
    /**
     * 水印颜色
     */
    private String color = "#C5CBCF";
    /**
     * 字体
     */
    private Font font = new Font("microsoft-yahei", Font.PLAIN, 16);
    /**
     * 水印图片宽度
     */
    private int width = 200;
    /**
     * 水印图片高度
     */
    private int height = 180;

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int getHeight() {
        return height;
    }

    public void setHeight(int height) {
        this.height = height;
    }

    public Font getFont() {
        return font;
    }

    public void setFont(Font font) {
        this.font = font;
    }
}
