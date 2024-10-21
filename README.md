# LITA--Excel-Class
This is where I document all Excel projects learnt from the Incubation Hub


[Project Overview](*project-overview)
[Excel functions](*excel-functions)
[Mathematical & Trigonometric Functions](*mathematics&trigonometrical-functions)
[Logical Functions](*logical-functions)
[Text Functions](*text-functions)
[Volatile Functions](*volatile-functions)
[Lookup & Reference Functions](*lookup&reference-functions)
[Statistical Functions](*statistical-functions)
[Conditional Functions](*conditional-functions)
[Date & Time Functions](*data&time-functions)
[Array Functions](*array-functions)
[CONDITIONAL FORMATTING](*conditional-formatting)
[REFERENCING](*referencing)
[Key Points](*key-points)


## Project 1: EXCEL FUNCTIONS

### Project Overview

Excel functions are powerful tools that simplify data analysis and calculations, making work easier and more efficient. This project dives into essential functions (NOT ALL), showing practical examples and real-life applications to help users build confidence and proficiency in handling various data tasks and decision-making.

## Excel functions
Excel functions are categorized based on their purpose and usage. Here is a breakdown of various types of Excel functions and their purposes:

### 1. Mathematical & Trigonometric Functions
   - **SUM**: Adds values in a range of cells.
   - **AVERAGE**: Calculates the mean of values in a range.
   - **ROUND**: Rounds a number to a specified number of digits.
   - **INT**: Rounds a number down to the nearest integer.
   - **MOD**: Returns the remainder after dividing two numbers.
   - **SIN, COS, TAN**: Calculate the sine, cosine, and tangent of an angle.

### 2. Logical Functions
   - **IF**: Checks whether a condition is met and returns one value if true, and another if false.
   - **AND**: Returns TRUE if all conditions are true, otherwise FALSE.
   - **OR**: Returns TRUE if any of the conditions are true.
   - **NOT**: Reverses the logic of its argument, returning FALSE for TRUE and vice versa.
   - **IFERROR**: Returns a specified value if an error is found in a formula.

### 3. Text Functions
   - **CONCATENATE** / **CONCAT** / **Textjoin**: Combines multiple strings into one.
   - **LEFT, RIGHT, MID, FIND, SEARCH**: Extracts characters from a string based on position.
   - **UPPER, LOWER, PROPER**: Converts text to uppercase, lowercase or proper case.
   - **LEN**: Returns the number of characters in a string.
   - **TRIM**: Removes extra spaces from a string, leaving single spaces between words.

### 4. Volatile Functions 
   Functions that recalculate anytime a change is made to the workbook, such as entering data, moving to another cell, or opening the workbook.
   - **RAND** / **RANDBETWEEN**: Returns a random number between two specified numbers
   - **RAND**: Returns a random number between 0 and 1. It generates a new random number every time Excel recalculates.
   - **TODAY**: Returns the current date.
   - **FORMULATEXT**: Displays the formula in a referenced cell as text.

### 5. Lookup & Reference Functions
   - **VLOOKUP**: Searches for a value in the first column of a range and returns a value in the same row from another column.
   - **HLOOKUP**: Similar to VLOOKUP but searches across the top row.
   - **INDEX**: Returns the value of a cell at the intersection of a row and column.
   - **MATCH**: Searches for a specified value in a range and returns its position.
   - **XLOOKUP**: An improved version of VLOOKUP/HLOOKUP, offering more flexibility.

### 6. Statistical Functions
   - **COUNT**: Counts the number of cells containing numbers in a range.
   - **COUNTA**: Counts the number of non-empty cells in a range.
   - **MIN, MAX**: Returns the smallest or largest value in a range.
   - **MEDIAN**: Finds the median value in a range.
   - **STDEV**: Calculates the standard deviation of a set of values.

### 7. Conditional Functions 
These functions perform calculations based on specific conditions or criteria, allowing you to aggregate data selectively.
   - **SUMIF**: Adds up values in a range that meet a specified condition.
   - **MINIF**: Finds the minimum value in a range that meets a condition. 
   - **MAXIF, MAXIFS**: Finds the maximum value in a range that meets a condition.

#### OTHER FUNCTIONS DISCOVERED:

### 1. Date & Time Functions
   - **TODAY**: Returns the current date.
   - **NOW**: Returns the current date and time.
   - **DATE**: Creates a date from year, month, and day values.
   - **YEAR, MONTH, DAY**: Extracts the year, month, or day from a date.
   - **DATEDIF**: Calculates the difference between two dates in days, months, or years.

### 2. Array Functions
     Functions that perform operations on a range of values (an array) rather than a single value, which enables calculations across multiple rows, columns, or ranges. They are useful for tasks like sorting, filtering, transposing, and manipulating data.
   - **WRAPROWS**: Takes a single row and splits into multiple rows based on a specified number of values per row.
   - **TRANSPOSE**: Converts a vertical range of cells to a horizontal range (or vice versa).
   - **FILTER**: Filters an array based on a specified condition.
   - **XLOOKUP**: An advanced lookup function that searches for a value in a range and returns the corresponding value from another range.
   - **TEXTSPLIT**: Splits a text string into an array based on a specified delimiter.


## CONDITIONAL FORMATTING
Conditional formatting in Excel is a feature that allows you to apply specific formatting—like colours, bolding, or font changes—to cells based on the values they contain. It is a tool for visually analyzing and highlighting important data, making it easier to spot trends, identify outliers, or flag specific entries.

### Examples of Conditional Formatting Uses
1. **Highlighting Duplicate Values**: identifies repeated entries in a dataset.
2. **Data Bars**: Displays bars within cells to visually represent the value's magnitude.
3. **Color Scales**: This applies a gradient of colours based on cell values, like shifting from red (low) to green (high) for easier comparison.
4. **Icon Sets**: Add icons (like arrows or flags) to indicate trends, performance, or categories based on cell values.
5. **Custom Formulas**: To use formulas to apply more specific formatting (e.g., colouring cells if a value is above the average).


### REFERENCING
1. **Absolute Referencing**: Keeps the cell reference constant, even when copied elsewhere. E.g $E$9
2. **Relative Referencing**: Changes when a formula is copied or dragged to another cell. E.g  E9
3. **Column Constant**: This is useful when you want only the column of the cell reference to remain constant. E.g $E9 
4. **Row Constant**: This is useful when you want only the row of the cell reference to remain constant. E.g E$9

#### Key Points:
1. To create a chart in your Pivot Table: **Alt + F1**
2. To open the format cell dialogue box: **Ctrl + 1**
3. To bring out the unique options in a column (Auto Filter): **Ctrl + Shift + L**
4. To lock a cell for Absolute Referencing: **F4**
5. Data Validation shortcut: **Alt + AVV**
6. Functional Argument Box / Highlight all cells: **Ctrl + A**
7. To hide a column: **Ctrl + O**
8. To freeze a customised random number: **Ctrl + ESV**
9. To convert data to the table: **Ctrl + T**
10. To make digits in thousands: **Ctrl + Shift + 1**
11. To separate first and last names into different columns (then follow the steps till achieved aim): **Ctrl + A +E**

