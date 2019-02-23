<table>
    <caption>${sheetName}</caption>
    <thead>
    <tr style="background-color: #6495ED">
        <th colspan="3" style="text-align: center;vertical-align: middle;font-weight: bold;font-size: 14px;">产品介绍</th>
    </tr>
    <tr>
        <#list titles as title>
            <th>${title}</th>
        </#list>
    </tr>
    </thead>
    <tbody>
    <#list data as item>
        <tr>
            <td>${item.category}</td>
            <td>${item.name}</td>
            <td>${item.count}</td>
        </tr>
    </#list>
    </tbody>
</table>