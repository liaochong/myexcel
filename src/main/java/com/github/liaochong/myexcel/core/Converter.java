package com.github.liaochong.myexcel.core;

/**
 * 映射提供者
 *
 * @author liaochong
 * @version 1.0
 */
public interface Converter<T, F> {
    /**
     * 转化原始数据为指定映射
     *
     * @param originalData 原始数据
     * @return 映射
     */
    F convert(T originalData);

}
