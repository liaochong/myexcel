package com.github.liaochong.myexcel.core.strategy;

import java.util.Objects;

/**
 * table 策略
 *
 * @author QingMings
 * @since v3.11.3
 */
public enum SheetStrategy {
    /**
     * 多个table生成在同一个sheet里
     */
    ONE_SHEET,
    /**
     * 每个table各生成一个sheet
     */
    MULTI_SHEET;

    public static boolean isOneSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, ONE_SHEET);
    }

    public static boolean isMultiSheet(SheetStrategy sheetStrategy) {
        return Objects.equals(sheetStrategy, MULTI_SHEET);
    }
}
