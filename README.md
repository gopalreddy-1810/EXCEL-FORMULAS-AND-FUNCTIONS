Excel Formulas & Functions Fundamentals

Project Overview

This project demonstrates commonly used Excel formulas and functions using the Sample Superstore Dataset. The main objective is to practice Excel functions used in everyday data analysis.

Dataset

Dataset: Sample Superstore Dataset

The dataset contains information such as:

- Order ID
- Order Date
- Customer Name
- Segment
- Category
- Sub-Category
- Sales
- Quantity
- Discount
- Profit
- Region

Functions Demonstrated

- VLOOKUP
- XLOOKUP
- INDEX/MATCH
- IF
- SUMIFS
- COUNTIFS
- LEFT
- RIGHT
- MID
- LEN
- TRIM
- UPPER
- LOWER
- CONCAT

Workbook Structure

1. Raw_Data – Original dataset
2. Lookup – VLOOKUP, XLOOKUP, and INDEX/MATCH examples
3. IF – Conditional logic examples
4. SUMIFS_COUNTIFS – SUMIFS and COUNTIFS examples
5. Text_Functions – Text manipulation examples
6. Summary_Notes – Explanation of formulas and their use cases

Edge Cases Tested

The formulas were tested with blank cells, invalid lookup values, and text-versus-number situations to understand how the formulas behave in different cases.

Interview Questions & Answers

1. What is the difference between SUMIF and SUMIFS?

SUMIF is used to calculate the sum of values based on a single condition.

Example:

=SUMIF(A:A,"Technology",B:B)

SUMIFS is used to calculate the sum based on multiple conditions.

Example:

=SUMIFS(C:C,A:A,"Technology",B:B,"West")

The main difference is that SUMIF works with one condition, while SUMIFS can work with multiple conditions.

2. How does XLOOKUP improve on VLOOKUP?

XLOOKUP is more flexible than VLOOKUP. It can search in any direction, does not require the lookup column to be the first column, and allows a custom value when no match is found.

Example:

=XLOOKUP(A2,F:F,G:G,"Not Found")

VLOOKUP generally searches from the first column of a selected table and returns a value from a column to its right.

Therefore, XLOOKUP provides more flexibility and easier error handling than VLOOKUP.

3. When would you use INDEX/MATCH instead of VLOOKUP?

INDEX/MATCH is useful when the lookup column is not on the left side of the return column or when more flexible lookup logic is required.

Example:

=INDEX(C:C,MATCH(A2,A:A,0))

INDEX/MATCH can look up values regardless of whether the return column is to the left or right of the lookup column.

Tools Used

- Microsoft Excel
- GitHub

Learning Outcomes

This project helped me practice:

- Data lookup functions
- Conditional calculations
- Text manipulation
- Data analysis
- Formula troubleshooting
- Handling basic data-quality issues

Conclusion

This project demonstrates the practical use of essential Excel formulas and functions for analyzing and working with sales data. It also helped me understand the appropriate use of lookup, conditional, aggregation, and text functions in real-world data analysis.

Author

Yeresi Venkata Ramagopal Reddy

B-TECH Data Science
