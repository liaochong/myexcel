package com.github.liaochong.html2excel.utils;

import java.util.function.IntSupplier;

/**
 * @author liaochong
 * @version 1.0
 */
public class TdUtils {

    public static int get(IntSupplier firstSupplier, IntSupplier secondSupplier) {
        int firstValue = firstSupplier.getAsInt();
        int secondValue = secondSupplier.getAsInt();
        return firstValue > 0 ? secondValue + firstValue - 1 : secondValue;
    }
}
