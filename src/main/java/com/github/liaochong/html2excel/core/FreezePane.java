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
public class FreezePane {

    /**
     * 从左到右需固定列数
     */
    int colSplit;

    /**
     * 从上到下需固定行数
     */
    int rowSplit;

}
