A dialog box for admins to email users about missing data fields in the PES (Part Evaluation Sheet)

## Screenshot
![PES Admin Functions screenshot][screenshot]

## Purpose
The Part Evaulation Sheet (PES) lists info about every part in the car (qty, weight, etc.), and it is vital this list has as much data as possible. Everything from the Cost Report during competition, to the shipping list passing through customs, relies on the PES. Tracking missing data manually would be incredibly time-consuming, since there are hundreds of parts, and many fields of data. Additionally, setting each field to 'required' when creating a part invites a host of issues, like garbage entries by creators for fields with unknown values (most items are entered before all information is available yet).

This tool was made to help notify each part creator of any missing information deemed necessary by a PES admin, in order to maximize the completion of info on each part and eliminate tedious workload(s).

## The code: how it works
Upon loading the dialog box from the toolbar on a Google Sheet (the PES), an admin would first select the subteams (ie. sheet names) and required fields (ie. column names); this forms the criteria to filter out missing part data. 

To aggregate the list of missing data, the sheet/column criteria is passed to `GetMissingPartData()` in Code.gs. This function checks for any blank cells with a column header matching one of the required fields (blank rows and filled cells are ignored), and creates an object for each part. for example:  
```
 {
            "sheetName": "Subteam Name",
            "partNumber": "46AB003",
            "partName": "Flux Capacitor",
            "missingData": [
                "Quantity",
                "Weight",
                "Gigawattage"
            ]
}
```

After looping through each row and making an array of part objects within each user, `GetMissingPartData()` returns an array of objects, for example: 
```
[
  {
    "user": "first.last@domain.com",
    "data": [{a part object}, {another part object}]
  },
  {
    "user": "first.last@domain.com",
    "data": [{a part object}]
  }
]
```

## Dialog usage
  1. Select sheet name(s) from list, to search for missing data
  2. Select column name(s) from list, to search for missing data
  3. (Optional) Edit email body
  4. Click "Generate Emails" button, to search for missing data and create email bodies
  5. Select recipient(s) from dropdown menu, to preview their email contents
  6. Click "Send emails" button, to send all emails

[screenshot]: screenshot.PNG
