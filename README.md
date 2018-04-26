# Personal.xlsb
A collection of utility functions of Excel plugin. 

![Ribbon](Image/Ribbon.PNG)

## Range Alignment Tool
A tool to align two ranges based on given columns. This impliment a full join function in any two given excel range.
To use this tool, import the RangeAlignmentTool.frm, RangeAlignmentTool.frx, and AlignTwoBlockBasedonGivenRange.bas

## Quick Format Tool
A tool quickly format a given range. If no range is selected, it will try to guess the given range. If the selected range contains infinit rows or columns, it will fail (to be fixed).
* It try to guess dates
* It try to guess percentage values but not very good at it

## Merge Selected Range Tool
It merges cells (row-wise) if they contains the same values. It is useful when copied data from database with the first several columns are sorted catagorical values.
