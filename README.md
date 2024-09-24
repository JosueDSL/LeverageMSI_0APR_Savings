# MSISavingsCalculator

The **MSISavingsCalculator** is a simple Python program that helps you calculate the interest savings when using an interest-free monthly payment plan (known as MSI, "meses sin intereses" or "0% APR") while keeping the financed money in a savings account that earns a specified annual interest rate. This strategy allows you to maximize your financial gain by earning interest on your savings while making monthly payments.

## Features

- Calculates the monthly payment amount.
- Computes interest earned each month on the remaining balance.
- Provides a month-by-month breakdown of your starting balance, interest earned, and ending balance.
- Displays the total interest earned over the entire payment period.
- Offers the option to view results in the terminal or save them to an Excel file with formatted tables.

## Requirements

- Python 3.x
- Required Python libraries (listed in `requirements.txt`):
  - `pandas`
  - `tabulate`
  - `openpyxl`

## Installation

1. **Clone the Repository or Download the Script**

   Save the `calculate.py` script to a directory on your computer.

2. **Set Up a Virtual Environment (Optional but Recommended)**

   It's good practice to use a virtual environment to manage dependencies.

   ```bash
   python -m venv venv
   ```

   Activate the virtual environment:

   - On Windows:

     ```bash
     venv\Scripts\activate
     ```

   - On macOS/Linux:

     ```bash
     source venv/bin/activate
     ```

3. **Install the Required Libraries**

   Make sure you have `pip` installed, then run:

   ```bash
   pip install -r requirements.txt
   ```

   This will install:

   - `pandas`: For data manipulation and analysis.
   - `tabulate`: For printing nicely formatted tables in the terminal.
   - `openpyxl`: For creating and formatting Excel files.

## Usage

Run the script using Python:

```bash
python calculate.py
```

The program will prompt you for the following inputs:

1. **Total Amount to Finance**

   - Enter the total purchase amount you wish to finance using MSI.
   - Example: `Enter the total amount to finance (in pesos): 20000`

2. **Annual Interest Rate**

   - Enter the annual interest rate of your savings account (as a percentage).
   - Example: `Enter the annual interest rate (e.g., 14 for 14%): 14`

3. **Number of Months to Finance**

   - Enter the number of months over which you'll be making payments.
   - Example: `Enter the number of months to finance (e.g., 12): 12`

4. **Output Preference**

   - Choose whether you'd like to view the results in the terminal or save them to an Excel file.
   - Example: `Would you like to save the results to an Excel file? (yes/no): no`

## Example Interaction

```plaintext
Enter the total amount to finance (in pesos): 20000
Enter the annual interest rate (e.g., 14 for 14%): 14
Enter the number of months to finance (e.g., 12): 12
Would you like to save the results to an Excel file? (yes/no): no

Summary:
+----------------------------------------+------------+
| Description                            | Value      |
|----------------------------------------+------------|
| Total Amount                           | $20,000.00 |
| Monthly Payment                        | $1,666.67  |
| Total Interest Earned                  | $1,645.77  |
| Interest If Untouched (Compounded)     | $2,986.83  |
| Interest If Untouched (Simple)         | $2,800.00  |
+----------------------------------------+------------+

Monthly Breakdown:
+---------+--------------------+-------------------+------------------+
|   Month |   Starting Balance |   Interest Earned |   Ending Balance |
|---------+--------------------+-------------------+------------------|
|       1 |           20000.00 |           237.81  |       18571.14   |
|       2 |           18571.14 |           217.41  |       17081.88   |
|       3 |           17081.88 |           199.00  |       15597.21   |
|       4 |           15597.21 |           179.81  |       14060.35   |
|       5 |           14060.35 |           161.28  |       12554.96   |
|       6 |           12554.96 |           143.85  |       11032.14   |
|       7 |           11032.14 |           126.76  |        946.22    |
|       8 |            946.22  |           109.27  |        789.27    |
|       9 |            789.27  |            91.91  |        631.51    |
|      10 |            631.51  |            74.58  |        472.75    |
|      11 |            472.75  |            56.70  |        313.02    |
|      12 |            313.02  |            38.50  |        149.15    |
+---------+--------------------+-------------------+------------------+
```

## Saving to Excel

If you choose to save the results to an Excel file by entering `yes` when prompted, the program will generate an Excel file named `MSI_Savings_Calculation.xlsx` in the same directory as the script.

The Excel file includes:

- A **Summary** section at the top with merged headers and formatted cells.
- A **Monthly Breakdown** table with headers and data formatted for readability.

## Notes

- **Data Accuracy**

  - The program assumes daily compounding interest and uses actual days per month for calculations.
  - Interest calculations are approximations and may vary slightly based on your financial institution's policies.

- **Error Handling**

  - The program includes input validation to ensure you enter positive numeric values.
  - If invalid input is detected, you'll be prompted to try again.

- **Dependencies**

  - Ensure all required libraries are installed using the `requirements.txt` file.
  - If you encounter issues with missing modules, double-check your virtual environment and installed packages.

## Contributing

Feel free to modify the script to suit your needs or contribute enhancements. If you add new features or fix bugs, consider sharing your improvements.

## License

This project is open-source and available for personal and educational use.

---

Enjoy maximizing your savings with the MSISavingsCalculator!