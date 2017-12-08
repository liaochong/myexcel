package com.github.liaochong.html2excel.core;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
class Td {
    /**
     * x坐标
     */
    int x;
    /**
     * y坐标
     */
    int y;
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
}
