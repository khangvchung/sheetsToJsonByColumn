# sheetsToJsonByColumn
Adds an "Export JSON" appscript that extracts google sheets data into JSON format by column rather than row.
<br><br>
**How to use**: **Go to Extensions > Apps Script, create a new script and paste the code in**
<br> <br>
**Requirements**: **In order for this to work you must go to view > freeze > 1 column**


Instead of extracting this with the headers at the top and each object is a row,
Name | number | email | address
--- | --- | --- | ---
Bob | 111-123-4567 | test@test.com | 1234 Jane St
Jane | 111-123-4567 | test@test.com | 1234 Jane St
Mary | 111-123-4567 | test@test.com | 1234 Jane St

It extracts from tables formatted like this where each object is the column.

| Name | Bob | Jane | Mary |
| --- | --- | --- | ---  |
| **number** | 111-123-4567 | 111-123-4567 | 111-123-4567 |
| **email** | test@test.com | test@test.com | test@test.com |
| **address** | 1234 Jane St | 1234 Jane St | 1234 Jane St |

The script also includes a "renameAndNestKey" function that renames or nests whatever key you want to rename.
