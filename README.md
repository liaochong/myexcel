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

- **使用动态页面模板引擎**：不用直接操作POI，使用自己常用的动态页面模板引擎生成html页面，学习成本更低；
- **可生成任意复杂表格**：本工具使用迭代单元格方式进行excel绘制，可生成任意复杂度excel；

Maven 依赖
------------------
```xml
<dependency>
    <groupId>com.github.liaochong</groupId>
    <artifactId>html2excel</artifactId>
    <version>0.0.1-beat</version>
</dependency>
```
