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
package com.github.liaochong.myexcel.core.style;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author liaochong
 * @version 1.0
 */
public class LinkDefaultCellStyle implements DefaultCellStyle {

    @Override
    public CellStyle supply(Workbook workbook) {
        CellStyle linkStyle = workbook.createCellStyle();
        Font linkFont = workbook.createFont();
        linkFont.setUnderline(Font.U_SINGLE);
        linkFont.setColor(IndexedColors.BLUE.getIndex());
        linkStyle.setFont(linkFont);
        linkStyle.setAlignment(HorizontalAlignment.CENTER);
        linkStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        linkStyle.setBorderBottom(BorderStyle.THIN);
        linkStyle.setBorderRight(BorderStyle.THIN);
        linkStyle.setBorderLeft(BorderStyle.THIN);
        linkStyle.setBorderTop(BorderStyle.THIN);
        return linkStyle;
    }
}
