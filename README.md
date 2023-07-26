# sheetsToJsonByColumn
Adds an "Export JSON" appscript that extracts google sheets data into JSON format by column rather than row.
<br><br>
Instead of extracting this with the headers at the top and each object is a row,
Name | number | email | address
--- | --- | --- | ---
Bob | 111-123-4567 | test@test.com | 1234 Jane St
Jane | 111-123-4568 | test@test.com | 1235 Jane St
Mary | 111-123-4569 | test@test.com | 1236 Jane St

It extracts from tables formatted like this where each object is the column.

| Name | Bob | Jane | Mary |
| --- | --- | --- | ---  |
| **number** | 111-123-4567 | 111-123-4568 | 111-123-4569 |
| **email** | test@test.com | test@test.com | test@test.com |
| **address** | 1234 Jane St | 1235 Jane St | 1236 Jane St |

The script also includes a "renameAndNestKey" function that renames or nests whatever key you want to rename.

**How to use**: 
1. Go to Extensions > Apps Script, create a new script and paste the code in
2. Go back to sheets, and reload page and a new option called "Export JSON" should appear at the top bar
3. **Requirement**: **In order for this to work you must go to view > freeze > 1 column**
4. Click "Export JSON for this sheet"
   ![alt text](images/Screenshot%201.png?raw=true)
6. A popup window should appear with the JSON objects
   ![alt text](images/Screenshot2.png?raw=true)
