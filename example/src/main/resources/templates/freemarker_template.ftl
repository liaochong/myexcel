<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Document</title>
</head>
<body>
<table>
    <tr>
        <th>标题一</th>
        <th>标题二</th>
        <th>标题三</th>
        <th>标题四</th>
    </tr>
    <tr>
        <td>${n_1}</td>
        <td>${n_2}</td>
        <td colspan="2">${n_3}</td>
    </tr>
    <tr>
        <td colspan="2">${n_4}</td>
        <td colspan="2">${n_5}</td>
    </tr>
    <tr>
        <td rowspan="2">${n_6}</td>
        <td colspan="2">${n_7}</td>
        <td>${n_8}</td>
    </tr>
    <tr>
        <td>${n_9}</td>
        <td>${n_10}</td>
        <td>${n_11}</td>
    </tr>
</table>
</body>
</html>