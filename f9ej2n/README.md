# Split bulk of text in rows and columns - Is there a faster way?

A solution to [Split bulk of text in rows and columns - Is there a faster way?](https://www.reddit.com/r/excel/comments/f9ej2n/split_bulk_of_text_in_rows_and_columns_is_there_a/) posted by u/RenardM

To use this, put the source file path and the sheet name respectively in the *file* cell and in the *sheet* cell of the *config* sheet. 

Then go to the *Reformat* sheet, right-click inside the green table and select *Refresh*.

&nbsp;

Requirements: Windows, Excel 2016 or higher, or Excel 2010/2013 with the Power Query add-in installed.

Note: this is based on my [PQ Template](https://github.com/tirlibibi17/excel-pq/tree/master/PQ%20Template) and uses the recently added RegexMatch function ^yay!