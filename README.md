# Expense Tracker
   #### Video Demo:  https://www.youtube.com/watch?v=dvoKs4l0zqM
   #### Description: 
   The Expense Tracker is a Python-based program designed to insert, view, manage, and export your personal expenses to Excel. 
It offers a menu-based interface that provides quick options for creating and managing expense records across multiple files.

## Features
- **Create New Expense Files**: Easily create yearly Excel files for expense tracking.
- **Monthly Expense Entries**: Insert your daily expenses with details like date, description, amount, and category.
- **Data View & Management**: View your monthly expenses, delete specific entries, and organize yearly totals by category.
- **Merge Yearly Expenses**: Aggregate expenses by category and get a summary of your yearly expenses.
- **Export**: Export the files to your own desktop for an easy access.

## Getting Started
Upon starting the program, you’ll see a menu with various options to manage and track expenses:

```text
Expense Tracker

1. Create New File
2. Insert Data
3. View Data
4. Erase Data
5. Merge Yearly Expenses
6. Export File
0. Exit

Select an option (0-6) from the menu.
```

### Menu Options
#### 1. Create New File
Create a new expense file for a specific year:
1. Enter the desired year (e.g., `2024`), and the program will create a new file named `Expenses2024.xlsx`.
   If a file already exists for that year, you will be prompted if you desire to replace it.


#### 2. Insert Data
Add expense records to a monthly sheet in an existing file:
1. Select a file from the available list of yearly expense files.
2. Choose a month (1-12) for data entry.
3. Enter expenses in an infinite loop, where each entry requires:
   - **Day** (Day of the expense)
   - **Description** (Expense details)
   - **Amount** (Cost of the expense)
   - **Category** (Expense category)

   Press `CTRL+C` to stop entering data and return to the main menu.

Each entry is automatically saved in the selected month's sheet, including:
   - 5 columns: Date, Description, Amount, Category, and a **Total** column.
   - The sum of the expenses for that month under the **Total** column.
   - Bold headers, filter options, and a randomly assigned header background color.


#### 3. View Data
View data entries for a specific month:
1. Choose a file and then a month with existing data.
2. The program will display all expenses from the selected month in a datatable format:

   ```text
   +-----------+---------------+----------+----------------+---------+
   | Day       | Description   |   Amount | Category       |   Total |
   +===========+===============+==========+================+=========+
   |  1/9/2024 | Rent          |     1200 | Housing        | 1288.26 |
   +-----------+---------------+----------+----------------+---------+
   | 18/9/2024 | Taxi          |    28.36 | Transportation |         |
   +-----------+---------------+----------+----------------+---------+
   | 30/9/2024 | Gym           |    59.90 | Health         |         |
   +-----------+---------------+----------+----------------+---------+
   ```


#### 4. Erase Data
Remove a specific row of data from a selected month:
1. Select a file and month.
2. The program will display all the entries from the selected month along with each row number:

   ```text
    +-------+-----------+---------------+--------+-----------------+---------+
    |       | Day       | Description   | Amount | Category        | Total   |
    +-------+-----------+---------------+--------+-----------------+---------+
    | Row 1 |  1/9/2024 | Rent          |   1200 | Housing         | 1288.26 |
    +-------+-----------+---------------+--------+-----------------+---------+
    | Row 2 | 18/9/2024 | Taxi          |  28.36 | Transportation |          |
    +-------+-----------+---------------+--------+-----------------+---------+
    | Row 3 | 30/9/2024 | Gym           |  59.90 | Health          |         |
    +-------+-----------+---------------+--------+-----------------+---------+
   ```
   
3. Enter the row number of the entry to delete (e.g., Row 1 to remove the first expense entry).


#### 5. Merge Yearly Expenses
Aggregate expenses by category for a selected year:
1. Select a file.
2. The program presents expenses by category, also displaying the number of occurrences, and the total amount per category. 
 A **Yearly Total** row shows the sum of all expenses for the year.

   **Example Format:**
   ```text
    +----------------+------------+----------+---------+
    | Category       | Quantity   | Amount   | Total   |
    +----------------+------------+----------+---------+
    | Health         | 2          | 119.8    | 252.04  |
    +----------------+------------+----------+---------+
    | Transportation | 2          | 55.65    |         |
    +----------------+------------+----------+---------+
    | Restaurants    | 1          | 76.59    |         |
    +----------------+------------+----------+---------+
   ```

3. The summary data is saved in a "Total" sheet, appended at the end of the file.


#### 6. Export File
Export any file to the user's desktop in `.xlsx` format, preserving all sheets and data.

## Requirements
- **Python**: Version 3 or higher
- **Libraries**: `openpyxl`, `tabulate`

Install openpyxl with:
```bash
pip install openpyxl
```
Install tabulate with:
```bash
pip install tabulate
```

## Usage
Run the program in a Python-supported IDE, then follow the on-screen prompts to interact with the menu and manage your expense records.

### Example Workflow
1. **Create a file** for a new year (Option 1).
2. **Insert monthly expenses** (Option 2) to keep track of your spending habits.
3. **View and validate entries** (Option 3) to review your expenses.
4. **Erase unwanted data** (Option 4) in case you have some expense to be erased.
5. **Generate a yearly report** (Option 5) to understand expense distribution.
6. **Export** your data (Option 6) to access your expenses outside the program.

## Goal
This program was built to simplify expense tracking by managing Excel files via terminal, allowing users to interact and update their data without directly editing Excel files.

---

Enjoy tracking your expenses more efficiently!
