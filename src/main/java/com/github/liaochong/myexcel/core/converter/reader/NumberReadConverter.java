/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core.converter.reader;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.function.Function;

/**
 * @param <R>
 * @author chd.y
 */
public class NumberReadConverter<R extends Number> extends AbstractReadConverter<R> {

    private Function<String, R> func;

    private NumberReadConverter(Function<String, R> func) {
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
     *
     * @param func 转换函数
     * @param <R>  目标类型
     * @return 转换器
     */
    public static <R extends Number> NumberReadConverter<R> of(Function<String, R> func) {
        return new NumberReadConverter<>(func);
    }

}
