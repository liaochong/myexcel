package com.github.liaochong.myexcel.core.converter.reader;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.function.Function;

/**
 * @author chd.y
 * @param <R>
 */
public class NumberReadConverter<R extends Number> extends AbstractReadConverter<R> {

    private Function<String, R> func;

    public NumberReadConverter(Function<String, R> func) {
        this.func = func;
    }

    @Override
    protected R doConvert(String v, Field field) {
        BigDecimal bigDecimal = new BigDecimal(v);
        String realValue = bigDecimal.toPlainString();
        return func.apply(realValue);

    }

    /**
     * 数字转换器
     * @param func
     * @param <R>
     * @return
     */
    public static <R extends Number> NumberReadConverter<R> of(Function<String, R> func) {
        return new NumberReadConverter<>(func);
    }

}
