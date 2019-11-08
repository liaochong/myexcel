<table>
    <caption>${sheetName}</caption>
    <thead>
    <tr style="background-color: #6495ED;height: 300px">
        <th colspan="3" style="text-align: center;vertical-align: middle;font-weight: bold;font-size: 14px;">产品介绍</th>
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
    <tr>
        <td colspan="3">fdssdfdsfhkadhkdajksdhajhdaskhakdhjkadf</td>
    </tr>
    <#--</#list>-->
    </tbody>
</table>