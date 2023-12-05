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
package com.github.liaochong.myexcel.core.parser;

import com.github.liaochong.myexcel.core.PromptContainer;
import com.github.liaochong.myexcel.utils.TdUtil;

import java.io.File;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collections;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public class Td {
    /**
     * 所在行
     */
    public int row;
    /**
     * 所在列
     */
    public int col;
    /**
     * 跨行数
     */
    public int rowSpan;
    /**
     * 跨列数
     */
    public int colSpan;
    /**
     * 内容
     */
    public String content;
    /**
     * 内容类型
     */
    public ContentTypeEnum tdContentType = ContentTypeEnum.STRING;
    /**
     * 是否为th
     */
    public boolean th;
    /**
     * 单元格样式
     */
    public Map<String, String> style = Collections.emptyMap();
    /**
     * 公式
     */
    public boolean formula;
    /**
     * 链接
     */
    public String link;
    /**
     * 文件
     */
    public File file;
    /**
     * 文件流
     */
    public InputStream fileIs;
    /**
     * 格式化
     */
    public String format;

    /**
     * 时间是常用对象，特殊化
     */
    public Date date;

    public LocalDate localDate;

    public LocalDateTime localDateTime;

    public List<Font> fonts;

    public PromptContainer promptContainer;
    /**
     * 斜线
     */
    public Slant slant;
    /**
     * 批注
     */
    public Comment comment;

    public Picture picture;

    public Td(int row, int col) {
        this.row = row;
        this.col = col;
    }

    public void setRowSpan(int rowSpan) {
        if (rowSpan < 2) {
            return;
        }
        this.rowSpan = rowSpan;
    }

    public void setColSpan(int colSpan) {
        if (colSpan < 2) {
            return;
        }
        this.colSpan = colSpan;
    }

    public int getRowBound() {
        return TdUtil.get(this.rowSpan, this.row);
    }

    public int getColBound() {
        return TdUtil.get(this.colSpan, this.col);
    }

    public Picture getPicture() {
        return picture;
    }

    public void setPicture(Picture picture) {
        this.picture = picture;
    }
}
