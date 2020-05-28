package com.github.liaochong.myexcel.core;

import java.util.List;

/**
 * @author liaochong
 * @version 1.0
 */
@FunctionalInterface
public interface ListSupplier<T> {

    /**
     * Gets a result.
     *
     * @return a result
     */
    List<T> getAsList();
}
