package com.github.liaochong.myexcel.core.templatehandler;

import com.github.liaochong.myexcel.core.parser.ParseConfig;
import com.github.liaochong.myexcel.core.parser.Table;

import java.util.List;
import java.util.Map;

/**
 * @author liaochong
 * @version 1.0
 */
public interface TemplateHandler {
    /**
     * 类路径模板
     *
     * @param path 类路径模板
     * @return TemplateHandler
     */
    TemplateHandler classpathTemplate(String path);

    /**
     * 文件路径模板
     *
     * @param dirPath  文件路径模板
     * @param fileName 模板名称
     * @return TemplateHandler
     */
    TemplateHandler fileTemplate(String dirPath, String fileName);

    /**
     * 获取模板字符流
     *
     * @param renderData 被渲染的数据
     * @param <E>        被渲染数据类型
     * @return 模板字符流
     */
    <E> String render(Map<String, E> renderData);

    <F> List<Table> render(Map<String, F> renderData, ParseConfig parseConfig) throws Exception;
}
