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

import com.github.liaochong.myexcel.utils.RegexpUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 数值读取转换
 *
 * @author chd.y
 * @since 2.3.1
 */
public class NumberReadConverter<R extends Number> extends AbstractReadConverter<R> {

    private static final Pattern PATTERN_ZERO = Pattern.compile("(.+)\\.0*");

    private Function<String, R> func;

    private NumberReadConverter(Function<String, R> func, boolean isInteger) {
        if (isInteger) {
            this.func = c -> {
                Matcher matcher = PATTERN_ZERO.matcher(c);
                boolean zeroSuffix = matcher.matches();
                return zeroSuffix ? func.apply(matcher.group(1)) : func.apply(c);
            };
        } else {
            this.func = func;
        }
    }

    @Override
    protected R doConvert(String v, Field field) {
        v = RegexpUtil.removeComma(v);
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
        return new NumberReadConverter<>(func, false);
    }

    /**
     * 数字转换器
     *
     * @param func      转换函数
     * @param <R>       目标类型
     * @param isInteger 是否为整数
     * @return 转换器
     */
    public static <R extends Number> NumberReadConverter<R> of(Function<String, R> func, boolean isInteger) {
        return new NumberReadConverter<>(func, isInteger);
    }
}
