package com.github.liaochong.myexcel.core.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author liaochong
 * @version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
@Documented
public @interface Prompt {

    /**
     * 标题
     *
     * @return 提示语
     */
    String title() default "title";

    /**
     * 提示
     *
     * @return 提示
     */
    String text() default "";
}
