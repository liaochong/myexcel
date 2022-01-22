package com.github.liaochong.myexcel.core;

/**
 * @author liaochong
 * @version 1.0
 */
@FunctionalInterface
interface CiConsumer<T, F, U> {

    void accept(T t, F f, U u);
}
