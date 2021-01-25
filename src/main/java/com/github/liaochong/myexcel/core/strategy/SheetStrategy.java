package com.github.liaochong.myexcel.core.strategy;

import java.util.Objects;

/**
 * table 策略
 * @author QingMings
 */
public enum SheetStrategy {
    /**
     * 多个table生成在同一个sheet里
     */
    OneSheet,
    /**
     * 每个table各生成一个sheet
     */
    MultiSheet;

    public static boolean isOneSheet(SheetStrategy sheetStrategy){
        return Objects.equals(sheetStrategy,OneSheet);
    }

    public static boolean isMultiSheet(SheetStrategy sheetStrategy){
        return Objects.equals(sheetStrategy,MultiSheet);
    }
}
