<table>
    <caption>${sheetName}</caption>
    <thead>
    <tr style="background-color: #6495ED;height: 100px">
        <th colspan="3"><span style="color: red">*</span><span style="font-weight: bold;">产品介绍</span></th>
    </tr>
    <tr>
        <#list titles as title>
            <th>${title}</th>
        </#list>
    </tr>
    </thead>
    <tbody>
    <#--<#list data as item>-->
    <#--<tr>-->
    <#--<td>${item.category}</td>-->
    <#--<td>${item.name}</td>-->
    <#--<td>${item.count}</td>-->
    <#--</tr>-->
    <tr>
        <td double>3,123.09</td>
        <td style="word-break: break-all">2<br/>676878>.~</td>
        <td>爱新觉罗·玄烨</td>
    </tr>
    <tr style="height: 100px;">
        <td>1</td>
        <td>2</td>
        <td>3</td>
    </tr>
    <tr style="background-color:red">
        <td colspan="3"><a href="http://www.baidu.com">百度链接</a></td>
    </tr>
    <#--</#list>-->
    </tbody>
</table>