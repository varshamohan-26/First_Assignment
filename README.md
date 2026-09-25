# First_Assignment

## 📊 Excel_Assignment_1_Data_Exploration

### 📝 Description

This assignment focuses on exploring and analyzing a product dataset using Microsoft Excel.

The main objective of this assignment is to practice Excel functions for calculating totals, counting products, finding averages, identifying minimum and maximum values, applying conditions, performing category-based calculations, and extracting text from Product IDs.

A separate Analysis sheet was created to perform the calculations while keeping the original dataset unchanged.

---

## 📂 Dataset

The dataset contains product information such as:

- Product ID
- Product Name
- Brand Name
- Price
- Quantity
- Category

# 🔍 Tasks Performed

## 1️⃣ SUM, COUNT and AVERAGE

### SUM Function

The SUM function was used to calculate the total price of all products.

Syntax:

=SUM(D2:D35)

Result:

Total Price = $10,100

---

### COUNT Function

The COUNT function was used to count the number of products based on the numerical values in the Price column.

Syntax:

=COUNT(D2:D35)

Result:

Number of Products = 34

---

### AVERAGE Function

The AVERAGE function was used to calculate the average price of all products.

Syntax:

=AVERAGE(D2:D35)

Result:

Average Price = $297.06

---

# 2️⃣ MIN and MAX Functions

### MIN Function

The MIN function was used to find the lowest product price.

Syntax:

=MIN(D2:D35)

Result:

Minimum Price = $30

---

### MAX Function

The MAX function was used to find the highest product price.

Syntax:

=MAX(D2:D35)

Result:

Maximum Price = $1,000

---

# 3️⃣ IF Function

The IF function was used to classify products according to their price.

If the product price is greater than or equal to $500, it is classified as "High Price".

If the product price is below $500, it is classified as "Standard Price".

Syntax:

=IF(D2>=500,"High Price","Standard Price")

---

# 4️⃣ SUMIF Function

The SUMIF function was used to calculate the total price of products belonging to the Electronics category.

Syntax:

=SUMIF(F2:F35,"Electronics",D2:D35)

Result:

Total Price of Electronics Products = $8,050

---

# 5️⃣ COUNTIF Function

The COUNTIF function was used to count the number of products with a price below $100.

Syntax:

=COUNTIF(D2:D35,"<100")

Result:

Number of Products Below $100 = 11

---

## LEFT Function

The LEFT function was used to extract the first two characters from the Product ID.

Syntax:

=LEFT(A2,2)

Example:

Product ID: 28-JAN-US

Result: 28

---

## RIGHT Function

The RIGHT function was used to extract the last two characters from the Product ID.

Syntax:

=RIGHT(A2,2)

Example:

Product ID: 28-JAN-US

Result: US

---

## MID Function

The MID function was used to extract the month from the Product ID.

Syntax:

=MID(A2,4,3)

Example:

Product ID: 28-JAN-US

Result: JAN

---

# 🎯 Learning Outcome

Through this assignment, I practiced using basic Excel functions for data exploration and analysis.

I learned how to:

- Calculate totals and averages
- Count products
- Find minimum and maximum values
- Apply conditions using the IF function
- Perform conditional calculations using SUMIF and COUNTIF
- Extract specific parts of text using LEFT, RIGHT and MID
- Organize analysis results in a separate worksheet
- Present analytical results clearly using Excel
