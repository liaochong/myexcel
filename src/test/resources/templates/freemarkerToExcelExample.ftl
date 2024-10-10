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
        <td prompt-title="提示" prompt-text="小提示1">爱新觉罗·玄烨</td>
    </tr>
    <tr style="height: 100px;">
        <td dropdownlist dropdownlist-name="省份">浙江,江西</td>
        <td dropdownlist dropdownlist-name="市区" dropdownlist-parent="省份">南昌,杭州,宁波</td>
        <td dropdownlist dropdownlist-parent="市区">上城区,下城区,弋阳,横峰</td>
    </tr>
    <tr style="background-color:red">
        <td colspan="3"><a href="http://www.baidu.com">百度链接</a></td>
    </tr>
    <#--</#list>-->
    </tbody>
</table>