An Apps Script dialog box to add, remove, and locate inventory items; bound to a Google Sheet.

## Screenshot
![inventory-tracking dialog screenshot][screenshot]

## Purpose
In GFR (Global Formula Racing, the student club at my university), we have a very large number of parts for our engine. It is important to keep track of all these parts (especially during engine rebuilds) in order to:

  - Reduce time wasted looking for parts
  - Save time/money spent buying (lost) parts we already purchased
  - Make engine rebuilds (a stressful task) easier and quicker
  - Reorder parts when empty
  - Allow new members to contribute with less confusion
  
This tool allows users to quickly add items to the inventory using a USB barcode scanner (for example, when a shipment of parts comes in, each one packaged in a barcoded plastic Honda bag), remove them using the same process, and lookup part info (without the dialog).

## The code: how it works
Coming soon&trade;

## Usage
Upon opening the Google Sheet, the toolbar is updated to show the custom "Add/Remove Parts" menu
  1. A user would first open the dialog through the toolbar
  2. Scan/enter a location code (barcode stickers are affixed to each location)
  3. Scan/enter a part barcode.  
    a. a subtotal list of parts is created automatically in the dialog
  4. (optional) Adjust quantity (using increment buttons, or by typing a number)
  5. (optional) Remove items from the subtotal list using the red "X" button on each row
  6. Verify subtotal list, then click "Submit" button at the bottom  
    a. Parts are automatically added/removed from the spreadsheet

## Attribution(s)
My first introduction to Google Apps Script came after watching **Lenny Cunningham**'s [Barcode Scanning with Chromebooks and Apps Script video on Youtube][lenny-link]. I owe a HUGE thank you to him and his work. Many nights were spent on Hangouts, holding my hand as I struggled through getting his project working on my own Sheet.

[screenshot]: screenshot.PNG
[lenny-link]: https://youtu.be/UON8jjI1GDc
