# 🧬 Excel-8

![Status](https://img.shields.io/badge/status-active-brightgreen.svg)
![Excel](https://img.shields.io/badge/Microsoft-Excel-blue.svg)

## ✨ Project Description

**Excel-8** is a guide to working with text in Microsoft Excel. Explore practical tips, instructions, and illustrations for manipulating text strings using functions like LEFT, RIGHT, MID, LEN, FIND, SUBSTITUTE, TRIM, Text To Columns, SEARCH, UPPER, LOWER, PROPER, EXACT, TEXT, CONCATENATE, TEXTJOIN, TEXTBEFORE, TEXTAFTER, and more.

> 📚 **Goal:** Help you master text operations in Excel—suitable for both beginners and advanced users!

> 🔗 **Note:** Many operations on text can be performed using Flash Fill described in [Excel-1 repository](https://github.com/Kuba27x/Excel-1).

---

## 📒 Table of Contents

- [Join Strings](#-join-strings)
- [LEFT](#-left)
- [RIGHT](#-right)
- [MID](#-mid)
- [LEN](#-len)
- [FIND & SEARCH](#-find--search)
- [SUBSTITUTE & REPLACE](#-substitute--replace)
- [TRIM & Counting Words](#-trim--counting-words)
- [Text to Columns](#-text-to-columns)
- [Change Case](#-change-case)
- [Compare Text](#-compare-text)
- [TEXT Function](#-text-function)
- [TEXTJOIN, TEXTBEFORE, TEXTAFTER](#-textjoin-textbefore-textafter)
- [Screenshots](#-screenshots)
- [Requirements](#-requirements)
- [Author](#-author)

---

## ➕ Join Strings

To join strings, use the `&` operator.

![Join](Screenshots/&.png)

Alternatively, use the `CONCATENATE` function.

---

## 🔡 LEFT

Extract the leftmost characters from a string:

![Left](Screenshots/Left.png)

---

## 🔢 RIGHT

Extract the rightmost characters from a string:

![Right](Screenshots/Right.png)

---

## 🎯 MID

Extract a substring starting from the middle of a string:

![Mid](Screenshots/Mid.png)

---

## 🔠 LEN

Get the length of a string:

![Len](Screenshots/Len.png)

---

## 🔍 FIND & SEARCH

Find the position of a substring:

- `FIND` (case-sensitive)
- `SEARCH` (case-insensitive, supports wildcards)

![Find](Screenshots/Find.png)

---

## 🔁 SUBSTITUTE & REPLACE

Replace existing text in a string:

- `SUBSTITUTE` replaces occurrences of text.
- `REPLACE` is useful when you know the position.

![Substitute](Screenshots/Substitute.png)

---

## 🧹 TRIM & Counting Words

- `TRIM` returns a string with only regular spaces.

![Trim](Screenshots/Trim.png)

- To get the length of the string without spaces:

![LenS](Screenshots/LenS.png)

- To count the number of words:

![Words](Screenshots/Words.png)

> ℹ️ **Tip:** Count the spaces in a cell and add 1 to get the total number of words.

---

## ✂️ Text to Columns

1. Select the range with full names.
2. On the Data tab, in the Data Tools group, click **Text to Columns**.
3. Choose **Delimited** and click **Next**.

![TextToColumns](Screenshots/TextToColumns.png)

4. Clear all checkboxes except for Comma and Space.

![Options](Screenshots/Options.png)

5. Click **Finish**.

**Result:**

![Result](Screenshots/Result.png)

---

## ↕️ Change Case

- Use `UPPER` to change text to uppercase:

![Upper](Screenshots/Upper.png)

- Use `LOWER` for lowercase.
- Use `PROPER` for capitalizing the first letter of each word:

![Proper](Screenshots/Proper.png)

---

## 🆚 Compare Text

- Use `EXACT` for case-sensitive comparison:

![Exact](Screenshots/Exact.png)

- Use `=F1=G1` for case-insensitive comparison:

![=](Screenshots/=.png)

---

## 🧾 TEXT Function

Format numbers when joining text and numbers:

![Text](Screenshots/Text.png)

> `#,##` adds commas to large numbers.

---

## 🔗 TEXTJOIN, TEXTBEFORE, TEXTAFTER

- `TEXTJOIN` (Excel 2016+) joins a range of strings using a delimiter, can ignore empty cells.

![TextJoin](Screenshots/TextJoin.png)

- In Excel 365, use `TEXTBEFORE` or `TEXTAFTER` to extract substrings:

![TextBefore](Screenshots/TextBefore.png)
![TextAfter](Screenshots/TextAfter.png)

---

## 📷 Screenshots

Find all screenshots in the `/Screenshots` folder.

---

## ℹ️ Requirements

- Microsoft Excel (recommended: 2016/2021/365 for modern functions)
- Windows OS for full functionality

---

## 👨‍💻 Author

Project and documentation by **Kuba27x**  
Repository: [Kuba27x/Excel-8](https://github.com/Kuba27x/Excel-8)

---
