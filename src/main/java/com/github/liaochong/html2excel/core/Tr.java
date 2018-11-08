package com.github.liaochong.html2excel.core;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
class Tr {
    /**
     * 索引
     */
    int index;
    /**
     * 行单元格
     */
    List<Td> tds = new ArrayList<>();
    /**
     * 行样式
     */
    Map<String, String> style;

    Tr(int index) {
        this.index = index;
    }
}
