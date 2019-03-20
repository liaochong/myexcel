package com.github.liaochong.myexcel.core.converter;

import java.lang.reflect.Field;

/**
 * @author liaochong
 * @version 1.0
 */
public interface ReadConverter {

    boolean convert(String content, Field field, Object obj) throws Exception;
}
