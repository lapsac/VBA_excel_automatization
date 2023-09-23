# VBA_excel_automatization
n the New Balance column, there is a FUNCTION called "n_balance" that calculates the new balance value in column M using the formula:

New Balance = Previous Balance + Invoice Amount - Paid.

In the Sibling Discount column, write a FUNCTION called "s_discount" to determine the discount from the Previous Balance column based on the number of children in the Enroled column. If the number of children is greater than 1, a 3% discount is applied for each child (e.g., 2 children - 6%, 3 children - 9%, etc.). The maximum discount that can be given is 15%.

In the Helper Discount column, there is a FUNCTION called "h_discount" that determines the discount based on the number of hours worked for the school. The condition is that if the number of hours worked in the Helper Hours column is greater than 15 and the employee status in the Emp? column is TRUE, then an additional discount of $250 is applied.

New Balance = Calculated fee - Sibling Discount - Helper Discount - Previous Balance + Paid Amount.
