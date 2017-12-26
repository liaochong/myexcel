# html2excel 文档
[![Build Status](https://travis-ci.org/liaochong/html2excel.svg?branch=master)](https://travis-ci.org/liaochong/html2excel)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.liaochong/html2excel/badge.svg)](https://maven-badges.herokuapp.com/maven-central/com.github.liaochong/html2excel)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)

html2excel是一款将html中表格（table）转化成excel的工具.

版本支持 | Support Version
------------------

- All version - only support for Java 8+

优点 | Advantages
------------------

- **零学习成本**：使用html作为模板，学习成本几乎为零；
- **屏蔽POI操作**：不直接操作过重的POI；
- **可生成任意复杂表格**：本工具使用迭代单元格方式进行excel绘制，可生成任意复杂度excel；
- **支持.XLS、.XLSX**：支持生成.xls、.xlsx后缀的excel；
- **支持多种模板引擎**：支持Freemarker、Beetl等常用模板引擎；

注意 | Attention
------------------
目前只支持模板文件存放在classpath下

Maven 依赖
------------------
```xml
<dependency>
    <groupId>com.github.liaochong</groupId>
    <artifactId>html2excel</artifactId>
    <version>0.0.1-beat</version>
</dependency>
```
