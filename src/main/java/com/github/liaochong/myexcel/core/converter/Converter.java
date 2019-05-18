package com.github.liaochong.myexcel.core.converter;

import java.lang.reflect.Field;

/**
 * 转换接口
 *
 * @author liaochong
 * @version 1.0
 */
public interface Converter<E, T> {

    /**
     * 转换
     *
     * @param obj   被转换对象
     * @param field 字段，提供额外信息
     * @return 转换结果
     */
    T convert(E obj, Field field);
}
