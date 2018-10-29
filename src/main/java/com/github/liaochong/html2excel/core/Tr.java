package com.github.liaochong.html2excel.core;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

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

    Map<String, String> style;

    Tr(int index) {
        this.index = index;
    }

    public void setStyle(Map<String, String> trStyle, Map<String, String> tableStyle) {
        if (Objects.isNull(trStyle) && Objects.isNull(tableStyle)) {
            return;
        }
        if (Objects.isNull(trStyle)) {
            this.style = tableStyle;
        } else {
            tableStyle.forEach(trStyle::putIfAbsent);
            this.style = trStyle;
        }
    }
}
