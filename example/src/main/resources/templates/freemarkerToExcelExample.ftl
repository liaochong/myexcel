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
    <caption>freemarker example</caption>
    <tbody>
    <#list data as item>
        <tr>
            <td>${item}</td>
        </tr>
    </#list>
    </tbody>
</table>
</body>
</html>