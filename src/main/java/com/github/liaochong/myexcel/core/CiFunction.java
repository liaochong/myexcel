package com.github.liaochong.myexcel.core;

/**
 * @author liaochong
 * @version 1.0
 */
@FunctionalInterface
interface CiFunction<T, F, R, U> {

    U apply(T t, F f, R r);
}
