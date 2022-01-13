# ExcelDataComparator
Excel Data Comparator for comparison between different sheets with different row orders and different row quantities.

As Data Integrity becomes more and more important, quickly identifying any data update, create and delete becomes a demand.

Excel is used in most office work, however the Excel standard function simply compares exactly the same cells in different sheets, which becomes inconvenient when the new sheet has different order, deleted or added rows. Hereby the Excel DataComparator project has been initiated.

This 1.1 version file provides flexible data comparison function.

It imports excel sheets from other files, and compare each cells of same rows in both old sheet, and new sheet then highlights the differences. The same rows in different sheets are identified by the Identify Column or columns, hereby the order of the rows in each sheet can be different, and the 2nd sheet is allow to delete some old rows, and add new rows, which are also highlighted in both sheets.

In some circumstances multiple columns may needed to distinguish rows, in this revision 1.1 Excel file, the only thing needed to do is concatenating the multiple column characters with comma “,”, and fill the string as the Identify Column value. A sample has been provided in the file.

