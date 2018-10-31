package com.github.liaochong.html2excel.core;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.util.Map;
import java.util.Objects;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
class Td {
    /**
     * 所在行
     */
    int row;
    /**
     * 所在列
     */
    int col;
    /**
     * 跨行数
     */
    int rowSpan;
    /**
     * 跨列数
     */
    int colSpan;
    /**
     * 内容
     */
    String content;
    /**
     * 是否为th
     */
    boolean th;
    /**
     * 单元格样式
     */
    Map<String, String> style;

    public void setStyle(Map<String, String> tdStyle, Map<String, String> trStyle) {
        if (Objects.isNull(trStyle) && Objects.isNull(tdStyle)) {
            return;
        }
        if (Objects.isNull(tdStyle)) {
            this.style = trStyle;
        } else {
            trStyle.forEach(tdStyle::putIfAbsent);
            this.style = tdStyle;
        }
    }

}
