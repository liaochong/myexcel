package com.github.liaochong.myexcel.core;

/**
 * @author liaochong
 * @version 1.0
 */
@FunctionalInterface
public interface AddTitleConsumer<T, E, F> {

    void accept(T t, E e, F f);
}
