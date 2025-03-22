# PythonGeeks Shop Management System

## Overview
The PythonGeeks Shop Management System is a simple GUI-based application designed to help manage products and sales records using an Excel database. The application provides functionality to add, delete, view products, and register sales transactions.

## Features
- **Add a Product**: Allows users to add a new product with a name, price, and date to the products database.
- **Delete a Product**: Enables users to remove a product by entering its name.
- **View Products**: Displays all products stored in the Excel database.
- **Register a Sale**: Records a sales transaction by inputting the customer name, product name, quantity, and date.

## Technologies Used
- Python
- Tkinter (for GUI)
- Pandas (for data manipulation and storage in Excel files)
- OpenPyXL (for handling Excel files)

## Installation
1. Ensure you have Python installed (Python 3 recommended).
2. Install the required Python libraries using:
   ```sh
   pip install pandas openpyxl
   ```
3. Download or clone the repository containing this project.
4. Run the script using:
   ```sh
   python shop_management.py
   ```

## How to Use
1. Launch the application by running the script.
2. Select the desired action from the main menu:
   - "Add a Product": Enter the required product details and click **ADD**.
   - "Delete a Product": Enter the product name and click **DELETE**.
   - "View Products": Displays a list of available products.
   - "Register a Sale": Enter customer and product details, then click **ADD**.
3. To exit the application, click **Quit** on any window.

## File Structure
- `products.xlsx`: Stores product details (Date, Product Name, Price).
- `sales.xlsx`: Stores sales transactions (Customer Name, Date, Product Name, Quantity, Price).
- `shop_management.py`: Main Python script containing the application logic.

## Notes
-products.xlsx` or `sales.xlsx` does not exist, the script automatically creates them.
- Data is stored persistently in Excel files for easy access and modification.


