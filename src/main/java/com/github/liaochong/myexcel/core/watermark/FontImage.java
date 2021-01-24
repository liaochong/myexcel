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
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;

/**
 * 文字图片
 *
 * @author liaochong
 * @version 1.0
 */
public class FontImage {

    public static BufferedImage createWatermarkImage(Watermark watermark) {
        BufferedImage image = new BufferedImage(watermark.getWidth(), watermark.getHeight(), BufferedImage.TYPE_INT_RGB);
        // 背景透明 开始
        Graphics2D g = image.createGraphics();
        image = g.getDeviceConfiguration().createCompatibleImage(watermark.getWidth(), watermark.getHeight(), Transparency.TRANSLUCENT);
        g.dispose();
        // 背景透明 结束
        g = image.createGraphics();
        // 设定画笔颜色
        g.setColor(new Color(Integer.parseInt(watermark.getColor().substring(1), 16)));
        g.setStroke(new BasicStroke(1));
        // 设置画笔字体
        g.setFont(watermark.getFont());
        // 设置字体平滑
        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        // 设定倾斜度
        g.rotate(-0.5, (double) image.getWidth() / 2, (double) image.getHeight() / 2);

        FontRenderContext context = g.getFontRenderContext();
        Rectangle2D bounds = watermark.getFont().getStringBounds(watermark.getText(), context);

        double x = (watermark.getWidth() - bounds.getWidth()) / 2;
        double y = (watermark.getHeight() - bounds.getHeight()) / 2;
        double ascent = -bounds.getY();
        double baseY = y + ascent;

        g.drawString(watermark.getText(), (int) x, (int) baseY);
        g.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
        // 释放画笔
        g.dispose();
        return image;
    }
}