package com.github.liaochong.html2excel.core;

import java.util.ArrayList;
import java.util.List;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

/**
 * @author liaochong
 * @version 1.0
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class Tr {
    /**
     * 索引
     */
    int index;
    /**
     * 行单元格
     */
    List<Td> tds = new ArrayList<>();

    public Tr(int index) {
        this.index = index;
    }
}
