package com.github.liaochong.myexcel.core.annotation;

import com.github.liaochong.myexcel.core.WorkbookType;
import com.github.liaochong.myexcel.core.strategy.WidthStrategy;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 替换ExcelTable
 *
 * @author liaochong
 * @version 1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
@Documented
public @interface ExcelModel {
    /**
     * 创建的excel包含所有字段
     *
     * @return true/false
     */
    boolean includeAllField() default true;

    /**
     * 是否忽略父类属性
     *
     * @return true/false
     */
    boolean excludeParent() default false;

    /**
     * 工作簿类型，.xls、.xlsx
     *
     * @return WorkbookType
     */
    WorkbookType workbookType() default WorkbookType.SXLSX;

    /**
     * 宽度策略
     *
     * @return WidthStrategy
     */
    WidthStrategy widthStrategy() default WidthStrategy.NO_AUTO;

    /**
     * sheeName
     *
     * @return sheeName
     */
    String sheetName() default "";

    /**
     * 是否使用字段名称作为标题
     *
     * @return true/false
     */
    boolean useFieldNameAsTitle() default false;

    /**
     * 为null时默认值
     *
     * @return 默认值
     */
    String defaultValue() default "";

    /**
     * 是否自动换行
     *
     * @return true/false
     */
    boolean wrapText() default true;

    /**
     * 是否过滤静态字段
     *
     * @return true/false
     */
    boolean ignoreStaticFields() default true;

    /**
     * 标题分离器
     *
     * @return 分离器
     */
    String titleSeparator() default "->";

    /**
     * 标题行高度
     *
     * @return 标题行高度
     */
    int titleRowHeight() default -1;

    /**
     * 普通行高度
     *
     * @return 普通行高度
     */
    int rowHeight() default -1;

    /**
     * 样式
     *
     * @return 样式
     */
    String[] style() default {};

    /**
     * 时间格式化
     *
     * @return 时间格式化
     */
    String dateFormat() default "";

    /**
     * 数值格式化
     *
     * @return 数值格式化
     */
    String decimalFormat() default "";
}
