Excel-8

Project Description

Excel-8 is a guide to working with text. Excel has many functions to offer when it comes to manipulating text strings. Here you'll find practical tips, instructions, and illustrations about those functions.

Table of Contents

Join Strings

To join strings, use the & operator.

![screenshot](Screenshots/&.png)

(Note: instead of using the & operator, you can use the CONCATENATE function.)

LEFT

To extract the leftmost characters from a string, use the LEFT function.

![screenshot](Screenshots/Left.png)

RIGHT

To extract the rightmost characters from a string, use the RIGHT function.

![screenshot](Screenshots/Right.png)

MID

To extract a substring, starting in the middle of a string, use the MID function.

![screenshot](Screenshots/Mid.png)

LEN

To get the length of a string, use the LEN function.

![screenshot](Screenshots/Len.png)

FIND

To find the position of a substring in a string, use the FIND function.

![screenshot](Screenshots/Find.png)

(Note: You can also use the SEARCH function. SEARCH function in Excel is case-insensitive and supports wildcards.

SUBSTITUTE

To replace existing text with new text in a string, use the SUBSTITUTE function.

![screenshot](Screenshots/Substitute.png)

(Note: If you only know the position of the text to be replaced, use the REPLACE function.)

Count words

TRIM

1. The TRIM function below returns a string with only regular spaces.

![screenshot](Screenshots/Trim.png)

To get the length of the string without spaces, use this formula.

![screenshot](Screenshots/LenS.png)

To get the number of words use this formula.

![screenshot](Screenshots/Words.png)

(Note: to count the total number of words in a cell, simply count the number of spaces and add 1 to this result. 1 space means 2 words, 2 spaces means 3 words)

Text to Columns

1. Select the range with full names.
2. On the Data tab, in the Data Tools group, click Text to Columns.
3. Choose Delimited and click Next.

![screenshot](Screenshots/TextToColumns.png)

4. Clear all the check boxes under Delimiters except for the Comma and Space check box.

![screenshot](Screenshots/Options.png)

5. Click Finish.
   
Result:

![screenshot](Screenshots/Result.png)

Change case

Use the UPPER function in Excel to change the case of text to uppercase.

![screenshot](Screenshots/Upper.png)

(Note: Use the LOWER function in Excel to change the case of text to lowercase.)

Use the PROPER function in Excel to change the first letter of each word to uppercase and all other letters to lowercase.

![screenshot](Screenshots/Proper.png)

Compare Text 

Use the EXACT function (case-sensitive).

![screenshot](Screenshots/Exact.png)

Use the formula =F1=G1 (case-insensitive).

![screenshot](Screenshots/=.png)

TEXT 

When joining text and a number, use the TEXT function in Excel to format that number.

![screenshot](Screenshots/Text.png)

(Note: #,## is used to add commas to large numbers.)

TEXTJOIN

The TEXTJOIN function in Excel 2016 or later joins a range of strings using a delimiter (first argument). It is very cool because it can ignore empty cells (if the second argument is set to TRUE).

![screenshot](Screenshots/TextJoin.png)

If you have Excel 365, use TEXTBEFORE or TEXTAFTER to extract substrings in Excel.

![screenshot](Screenshots/TextBefore.png)
![screenshot](Screenshots/TextAfter.png)

(Note: many operations on text can be performed using Flash Fill described in Excel-1 repository)

