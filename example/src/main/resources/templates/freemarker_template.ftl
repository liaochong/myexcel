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
<table border="1">
    <!-- Row A -->
    <tr>
        <td colspan="3">1</td>
        <td>4</td>
        <td rowspan="3">5</td>
    </tr>
    <!-- Row B -->
    <tr>
        <td>1</td>
        <td rowspan="2">2</td>
        <td>3</td>
        <td>4</td>
    </tr>
    <!-- Row C -->
    <tr>
        <td rowspan="2">1</td>
        <td>3</td>
        <td>4</td>
    </tr>
    <!-- Row D -->
    <tr>
        <td colspan="3">2</td>
        <td>5</td>
    </tr>
    <!-- Row E -->
    <tr>
        <td>1</td>
        <td>2</td>
        <td colspan="3">3</td>
    </tr>
</table>

<table border="1">

    <!-- First row -->

    <tr>
        <td>1</td>
        <td colspan="2">2 and 3</td>
        <td>4</td>
    </tr>

    <!-- Second row -->

    <tr>
        <td rowspan="3">5, 9 and 13</td>
        <td>6</td>
        <td>7</td>
        <td>8</td>
    </tr>

    <!-- Third row -->

    <tr>
        <td>10</td>
        <td>11</td>
        <td>12</td>
    </tr>

    <!-- Fourth row -->

    <tr>
        <td colspan="3">14, 15 and 16</td>
    </tr>

</table>

<table>
    <col>
    <colgroup span="2"></colgroup>
    <colgroup span="2"></colgroup>
    <tr>
        <td rowspan="2"></td>
        <th colspan="2" scope="colgroup">Mars</th>
        <th colspan="2" scope="colgroup">Venus</th>
    </tr>
    <tr>
        <th scope="col">Produced</th>
        <th scope="col">Sold</th>
        <th scope="col">Produced</th>
        <th scope="col">Sold</th>
    </tr>
    <tr>
        <th scope="row">Teddy Bears</th>
        <td>50,000</td>
        <td>30,000</td>
        <td>100,000</td>
        <td>80,000</td>
    </tr>
    <tr>
        <th scope="row">Board Games</th>
        <td>10,000</td>
        <td>5,000</td>
        <td>12,000</td>
        <td>9,000</td>
    </tr>
</table>
<table id="ctbl-01">
    <caption>
        Imported and domestic orange and apple prices
        in Sydney and Melbourne
    </caption>
    <thead>
    <tr>
        <td></td>
        <th colspan="2" id="imported">Imported</th>
        <th colspan="2" id="domestic">Domestic</th>
    </tr>
    <tr>
        <td></td>
        <th id="oranges-imp">Oranges</th>
        <th id="apples-imp">Apples</th>
        <th id="oranges-dom">Oranges</th>
        <th id="apples-dom">Apples</th>
    </tr>
    </thead>
    <tbody>
    <tr>
        <th id="sydney" colspan="5">Sydney</th>
    </tr>
    <tr>
        <th headers="sydney" id="wholesale-sydney">Wholesale</th>
        <td headers="imported oranges-imp sydney wholesale-sydney">$1.00</td>
        <td headers="imported apples-imp sydney wholesale-sydney">$1.25</td>
        <td headers="domestic oranges-dom sydney wholesale-sydney">$1.20</td>
        <td headers="domestic apples-dom sydney wholesale-sydney">$1.00</td>
    </tr>
    <tr>
        <th headers="sydney" id="retail-sydney">Retail</th>
        <td headers="imported oranges-imp sydney retail-sydney">$2.00</td>
        <td headers="imported apples-imp sydney retail-sydney">$3.00</td>
        <td headers="domestic oranges-dom sydney retail-sydney">$1.80</td>
        <td headers="domestic apples-dom sydney retail-sydney">$1.60</td>
    </tr>
    <tr>
        <th id="melbourne" colspan="5">Melbourne</th>
    </tr>
    <tr>
        <th headers="melbourne" id="wholesale-melbourne">Wholesale</th>
        <td headers="imported oranges-imp melbourne wholesale-melbourne">$1.20</td>
        <td headers="imported apples-imp melbourne wholesale-melbourne">$1.30</td>
        <td headers="domestic oranges-dom melbourne wholesale-melbourne">$1.00</td>
        <td headers="domestic apples-dom melbourne wholesale-melbourne">$0.80</td>
    </tr>
    <tr>
        <th headers="melbourne" id="retail-melbourne">Retail</th>
        <td headers="imported oranges-imp melbourne retail-melbourne">$1.60</td>
        <td headers="imported apples-imp melbourne retail-melbourne">$2.00</td>
        <td headers="domestic oranges-dom melbourne retail-melbourne">$2.00</td>
        <td headers="domestic apples-dom melbourne retail-melbourne">$1.50</td>
    </tr>
    </tbody>
</table>

<table>
    <caption>
        Prices of Brass and Steel nuts and bolts
    </caption>
    <colgroup>
    </colgroup>
    <colgroup span="2">
    </colgroup>
    <colgroup span="2">
    </colgroup>
    <thead>
    <tr>
        <td></td>
        <th scope="colgroup" colspan="2">Brass</th>
        <th scope="colgroup" colspan="2">Steel</th>
    </tr>
    <tr>
        <td></td>
        <th scope="col">Bolts</th>
        <th scope="col">Nuts</th>
        <th scope="col">Bolts</th>
        <th scope="col">Nuts</th>
    </tr>
    </thead>
    <tbody>
    <tr>
        <th scope="rowgroup" colspan="5">10cm</th>
    </tr>
    <tr>
        <th scope="row">Wholesale</th>
        <td>$1.00</td>
        <td>$1.25</td>
        <td>$1.20</td>
        <td>$1.00</td>
    </tr>
    <tr>
        <th scope="row">Retail</th>
        <td>$2.00</td>
        <td>$3.00</td>
        <td>$1.80</td>
        <td>$1.60</td>
    </tr>
    <tr>
        <th scope="rowgroup" colspan="5">20cm</th>
    </tr>
    <tr>
        <th scope="row">Wholesale</th>
        <td>$1.20</td>
        <td>$1.30</td>
        <td>$1.00</td>
        <td>$0.80</td>
    </tr>
    <tr>
        <th scope="row">Retail</th>
        <td>$1.60</td>
        <td>$2.00</td>
        <td>$2.00</td>
        <td>$1.50</td>
    </tr>
    </tbody>
</table>

<table>
    <caption>
        Imported and domestic cherry and apricot prices in Perth and Adelaide
    </caption>
    <thead>
    <tr>
        <td></td>
        <th id="imp" colspan="3">Imported</th>
        <th id="dom" colspan="3">Domestic</th>
    </tr>
    <tr>
        <td></td>
        <th headers="imp" id="imp-apr">Apricots</th>
        <th headers="imp" id="imp-che" colspan="2">Cherries</th>
        <th headers="dom" id="dom-apr">Apricots</th>
        <th headers="dom" id="dom-che" colspan="2">Cherries</th>
    </tr>
    <tr>
        <td></td>
        <td></td>
        <th headers="imp imp-che" id="imp-che-agrade">A Grade</th>
        <th headers="imp imp-che" id="imp-che-bgrade">B Grade</th>
        <td></td>
        <th headers="dom dom-che" id="dom-che-agrade">A Grade</th>
        <th headers="dom dom-che" id="dom-che-bgrade">B Grade</th>
    </tr>
    </thead>
    <tbody>
    <tr>
        <th id="perth" colspan="7">Perth</th>
    </tr>
    <tr>
        <th headers="perth" id="perth-wholesale">Wholesale</th>
        <td headers="imp imp-apr perth perth-wholesale">$1.00</td>
        <td headers="imp imp-che imp-che-agrade perth perth-wholesale">$9.00</td>
        <td headers="imp imp-che imp-che-bgrade perth perth-wholesale">$6.00</td>
        <td headers="dom dom-apr perth perth-wholesale">$1.20</td>
        <td headers="dom dom-che dom-che-agrade perth perth-wholesale">$13.00</td>
        <td headers="dom dom-che dom-che-bgrade perth perth-wholesale">$9.00</td>
    </tr>
    <tr>
        <th headers="perth" id="perth-retail">Retail</th>
        <td headers="imp imp-apr perth perth-retail">$2.00</td>
        <td headers="imp imp-che imp-che-agrade perth perth-retail">$12.00</td>
        <td headers="imp imp-che imp-che-bgrade perth perth-retail">$8.00</td>
        <td headers="dom dom-apr perth perth-retail">$1.80</td>
        <td headers="dom dom-che dom-che-agrade perth perth-retail">$16.00</td>
        <td headers="dom dom-che dom-che-bgrade perth perth-retail">$12.50</td>
    </tr>
    <tr>
        <th id="adelaide" colspan="7">Adelaide</th>
    </tr>
    <tr>
        <th id="adelaide-wholesale">Wholesale</th>
        <td headers="imp imp-apr adelaide adelaide-wholesale">$1.20</td>
        <td headers="imp imp-che imp-che-agrade adelaide adelaide-wholesale">N/A</td>
        <td headers="imp imp-che imp-che-bgrade adelaide adelaide-wholesale">$7.00</td>
        <td headers="dom dom-apr adelaide adelaide-wholesale">$1.00</td>
        <td headers="dom dom-che dom-che-agrade adelaide adelaide-wholesale">$11.00</td>
        <td headers="dom dom-che dom-che-bgrade adelaide adelaide-wholesale">$6.00</td>
    </tr>
    <tr>
        <th id="adelaide-retail">Retail</th>
        <td headers="imp imp-apr adelaide adelaide-retail">$1.60</td>
        <td headers="imp imp-che imp-che-agrade adelaide adelaide-retail">N/A</td>
        <td headers="imp imp-che imp-che-bgrade adelaide adelaide-retail">$11.00</td>
        <td headers="dom dom-apr adelaide adelaide-retail">$2.00</td>
        <td headers="dom dom-che dom-che-agrade adelaide adelaide-retail">$13.00</td>
        <td headers="dom dom-che dom-che-bgrade adelaide adelaide-retail">$10.00</td>
    </tr>
    </tbody>
</table>
</body>
</html>
