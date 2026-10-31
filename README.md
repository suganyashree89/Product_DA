
# 📊 Product Database Analysis Using Microsoft Excel

## 📌 Project Overview

This project focuses on Product Database Analysis using Microsoft Excel. The dataset contains product information such as Product Name, Brand Name, Price, Quantity, Category, and Price Range.

The project uses Excel Logical Functions and Text Processing Functions to analyze product data, classify prices, and extract meaningful information from product codes.

The main objective is to improve Excel skills in data cleaning, data processing, logical operations, and product data analysis.

## 🎯 Project Objectives

- Analyze product information using Microsoft Excel.
- Apply logical functions to classify product prices.
- Extract day, country code, and month from product codes.
- Process text data using Excel text functions.
- Organize product information into meaningful categories.
- Improve data analysis and Excel formula skills.

## 📂 Dataset Description

The dataset contains the following columns:

| Column Name | Description |
|---|---|
| Product Code | Contains day, month, and country code |
| Product Name | Name of the product |
| Brand Name | Product brand |
| Price ($) | Product price in dollars |
| Quantity | Available product quantity |
| Category | Product category |
| Price Range | Product price classification |
| Day | Extracted day from product code |
| Country Code | Extracted country code |
| Month With Code | Extracted month and country code |
| Month | Extracted month |

## 🛠️ Tools & Technologies

- Microsoft Excel
- Logical Functions
- Text Processing Functions
- Data Cleaning
- Data Analysis
- Excel Formulas

## 🧮 Excel Functions Used

### 1. Logical Functions

**IF Function**

Used to classify products based on their prices.

```excel
=IF(D2>=500,"High Price","Standard Price")
```

**AND Function**

Used to check multiple conditions simultaneously.

```excel
=IF(AND(D2>=500,E2>=20),"High Value Stock","Regular Stock")
```

### 2. Text Processing Functions

**LEFT Function**

Extracts the first characters from the Product Code.

```excel
=LEFT(A2,2)
```

**RIGHT Function**

Extracts the country code from the Product Code.

```excel
=RIGHT(A2,2)
```

**MID Function**

Extracts the month from the Product Code.

```excel
=MID(A2,4,3)
```


## 📊 Sample Dataset

| Product Name | Brand | Price ($) | Quantity | Category | Price Range |
|---|---|---:|---:|---|---|
| Laptop | Dell | 1000 | 30 | Electronics | High Price |
| Sneakers | Nike | 80 | 15 | Fashion | Standard Price |
| Coffee Maker | Keurig | 130 | 40 | Kitchen | Standard Price |
| Smartphone | Samsung | 900 | 25 | Electronics | High Price |
| Backpack | North Face | 70 | 20 | Outdoor | Standard Price |

## 📈 Key Analysis

- Identified high-price and standard-price products.
- Extracted day information from product codes.
- Identified country codes.
- Extracted month information.
- Combined text values using Excel functions.
- Organized product data for easier analysis.

## 🚀 Project Benefits

- Improves knowledge of Excel logical functions.
- Develops text manipulation skills.
- Helps understand product price classification.
- Simplifies data extraction and transformation.
- Improves practical data analysis skills.

## 🔮 Future Improvements

- Create an interactive Excel dashboard.
- Add Pivot Tables for category-wise analysis.
- Visualize product prices using charts.
- Perform brand-wise product analysis.
- Analyze total inventory value.
- Add more logical and text functions.

## 👨‍💻 Author

**suganyashree**

GitHub: https://github.com/suganyashree89



