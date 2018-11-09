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
<table class="complex" cellpadding="2" cellspacing="0" summary="Purchase Receipt Information"
       style="text-align: center">


    <tbody>
    <tr style="background-color: yellowgreen">
        <th class="sect" colspan="10" id="ri">Requester Information</th>
    </tr>

    <tr style="background-color: red">
        <th colspan="1" id="name" class="label">Name</th>
        <td colspan="7" headers="name" style="background-color: yellow; border-bottom-style: dashed">Jon Smith</td>
        <th colspan="1" id="date" class="label">Date</th>
        <td colspan="1" headers="date">9/9/2005</td>
    </tr>

    <tr>
        <th colspan="1" id="dept" class="label">Department</th>
        <td colspan="9" headers="ri dept">Customer Service</td>
    </tr>

    <tr style="font-weight: bold;font-size: 14px;">
        <th colspan="1" id="remail" class="label">E-mail</th>
        <td colspan="4" headers="ri remail">jon.smith@stateu.edu</td>
        <th colspan="1" id="rphone" class="label">Phone</th>
        <td colspan="4" headers="ri rphone">(123) 456-7890</td>
    </tr>

    <tr>
        <th class="sect" colspan="10" id="si" style="font-style: italic;">Supplier Information</th>
    </tr>

    <tr>
        <th colspan="1" id="cn" class="label">Name</th>
        <td colspan="9" headers="si cn">Acme Computer Supplies</td>
    </tr>

    <tr>
        <th colspan="1" id="ad1" class="label">Address 1</th>
        <td colspan="9" headers="si ad1">1234 Business Way</td>
    </tr>

    <tr>
        <th colspan="1" id="ad2" class="label">Address 2</th>
        <td colspan="9" headers="si ad2"></td>
    </tr>

    <tr>
        <th colspan="1" id="city" class="label">City</th>
        <td colspan="3" headers="si city">Somewhere</td>
        <th colspan="1" id="state" class="label">State</th>
        <td colspan="1" headers="si state"><abbr title="Illinois">IL</abbr></td>
        <th colspan="3" id="zip" class="label">Zip Code</th>
        <td colspan="1" headers="si zip">61820</td>
    </tr>

    <tr>
        <th colspan="1" class="label" id="semail">E-mail</th>
        <td colspan="4" headers="si semail">order@acmesupply.edu</td>
        <th colspan="1" id="sweb" class="label">Web</th>
        <td colspan="4" headers="si sweb"><a href="http://www.acmesupply.edu/">www.acmesupply.edu</a></td>
    </tr>

    <tr>
        <th colspan="1" id="sphone" class="label">Phone</th>
        <td colspan="4" headers="si sphone">(123) 222-3344</td>
        <th colspan="1" id="sfax" class="label">Fax</th>
        <td colspan="4" headers="si sfax">(123) 222-3355</td>
    </tr>

    <tr>
        <th class="sect" colspan="10" id="items">Items</th>
    </tr>

    <tr>
        <th class="title" colspan="1" id="itemnum">Item #</th>
        <th colspan="1" class="title" id="part">Part Number</th>
        <th colspan="5" class="title" id="desc">Description</th>
        <th colspan="1" id="qty" class="title">Quantity</th>
        <th colspan="1" id="unit" class="title">Unit Cost</th>
        <th colspan="1" id="totalcost" class="title">Total Cost</th>
    </tr>

    <tr>
        <th class="title" colspan="1" id="num1">1</th>
        <td class="part" colspan="1" headers="part  num1">197-540501</td>
        <td class="part" colspan="5" headers="desc  num1">Toner Cartridge for LaserJet 4100 Series 6000 Page Duty
            Cycle
        </td>
        <td class="qty" colspan="1" headers="qty   num1">2</td>
        <td class="cost" colspan="1" headers="unit  num1">$29.95</td>
        <td class="cost" colspan="1" headers="totalcost num1">$59.90</td>
    </tr>

    <tr>
        <th class="title" colspan="1" id="num2">2</th>
        <td class="part" colspan="1" headers="part  num2">555-013097</td>
        <td class="part" colspan="5" headers="desc  num2">Ink Cartridge color for Epson Stylus C62</td>
        <td class="qty" colspan="1" headers="qty   num2">3</td>
        <td class="cost" colspan="1" headers="unit  num2">$27.00</td>
        <td class="cost" colspan="1" headers="totalcost num2">$81.00</td>
    </tr>

    <tr>
        <th class="title" colspan="1" id="num3">3</th>
        <td class="part" colspan="1" headers="part  num3">555-0167213</td>
        <td class="part" colspan="5" headers="desc  num3">Parallel Printer Cable 10ft</td>
        <td class="qty" colspan="1" headers="qty   num3">1</td>
        <td class="cost" colspan="1" headers="unit  num3">$6.00</td>
        <td class="cost" colspan="1" headers="totalcost num3">$6.00</td>
    </tr>

    <tr>
        <th class="sub" colspan="9" id="subtotal">Sub-Total</th>
        <td class="cost" colspan="1" headers="subtotal">$146.90</td>
    </tr>

    <tr>
        <th class="sub" colspan="9" id="tax">Tax</th>
        <td class="cost" colspan="1" headers="tax">$10.28</td>
    </tr>

    <tr>
        <th class="sub" colspan="9" id="ship">Shipping</th>
        <td class="cost" colspan="1" headers="ship">$15.80</td>
    </tr>

    <tr>
        <th colspan="9" id="totalall" class="total">Total</th>
        <td class="total" colspan="1" headers="totalall">$172.98</td>
    </tr>

    <tr>
        <th colspan="10" id="approval" class="sect">Approval</th>
    </tr>

    <tr>
        <th class="approval" colspan="1" id="appperson">Person</th>
        <td class="approval" colspan="7" headers="approval appperson">Sara Johnson</td>
        <th class="approval" colspan="1" id="appdate">Date</th>
        <td class="approval" colspan="1" headers="approval appdate">9/7/2005</td>
    </tr>

    </tbody>

</table>
</body>
</html>
