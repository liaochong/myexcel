<!doctype html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
</head>
<body>
<table>
    <caption>${sheetName}</caption>
    <thead>
    <tr>
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
</body>
</html>