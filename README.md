# Excel Barcode Scanner Guide
Excel technique that allows you to quickly scan a barcode (or QR code) to find a row and manually input a value associated with that row. In our lab we are using this to scan plants and record measurements while in the greenhouse.

# Steps
1. Obtain a barcode or QR code scanner which can be configured to add a custom suffix upon a scan. A lot of scanners have this functionality for carriage returns (`Enter`), but in our case we need it to be able to have `Tab` as the suffix. 
    1. We happened to go with this scanner: https://www.amazon.com/REALINN-Wireless-Rechargeable-Automatic-Warehouse/dp/B07W5Q48M5/. Refer to Page 16 of the [manual](https://github.com/Salk-Harnessing-Plants-Initiative/excel-barcode-scanner-guide/blob/main/RL6200W%20Manual%20V1.0.pdf) to configure the scanner to give `Tab` instead of carriage return and the end of each scan.

# Credit
Credit for the original solution goes to [steveingmc on Reddit](https://www.reddit.com/r/excel/comments/2p1xuf/locating_a_cell_or_item_using_a_bar_code_scanner/cmsxri4?utm_source=share&utm_medium=web2x&context=3):

> To start off we did do a a small bit of VBA to search the excel sheet once the Barcode had scanned, 
but the users didn't like it. They found it faster using the second way described below.
>
> We do something very similar at my work, generally used for stock taking. Here's how we do it.
>
> On the first sheet called [SystemData] we dump from our main system, 
>the stock list with the following data in Columns A to C
>
>Stock ID, Description and QTY
>
>On the second sheet named [Scan Data] The barcode scanner is programmed to scan the Item Barcode, 
>add the scanned data to column A and then send the TAB command moving the selected cell to column B, 
>which the User then enters the stock quantity.
>
>On this same sheet [Scan Data] is a preset Vlookup that references the [System Data] sheet columns A to C 
>and returns the Stock Description and Quantity. This is just a safe guard really and allows the user to see 
>that the scanned item appears in the master system list etc.
>
>Now back on the [System Data] sheet in column D we have another vlookup, looking back at columns A & B on sheet [Scan Data]
>This lookup looks for the new Quantity count and returns the value. If it cant find a lookup match, 
>i.e. returns a #N\A we know that either theres no stock or we missed that item when scanning, allowing us to go check.
>
>The final column on the [System Data] sheet is an IF statement comparing the System Stock Qty and the new Human entered value. 
>If they differ, the new count is returned, if they are the same, keep that value.
>
>There are a few condition formats, to highlight massive variances or where we have more counted stock 
>that the system thinks we should have.

