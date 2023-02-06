import sys
import os
from datetime import date
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    if len(sys.argv) < 2:
        print("error. no path provided.")
        sys.exit(1)

    # Check whether provide parameter is valid path of file
    if os.path.isfile(sys.argv[1]):
        return sys.argv[1]
    else:
        print("error. target of path is not a file.")
        sys.exit(1)

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    csv_parent_dir = os.path.dirname(sales_csv)

    # Determine the name and path of the directory to hold the order data files
    todays_date = date.today().isoformat()
    orders_dir_name = f'Order_{todays_date}'
    orders_dir_path = os.path.join(csv_parent_dir,orders_dir_name)

    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir_path):
        os.makedirs(orders_dir_path)

    return orders_dir_path

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_df = pd.read_csv(sales_csv)

    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, "TOTAL PRICE", sales_df['ITEM QUANTITY'] * sales_df["ITEM PRICE"])

    # Remove columns from the DataFrame that are not needed
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)

    # Group the rows in the DataFrame by order ID
    for order_id, order_df in sales_df.groupby('ORDER ID'):
    # For each order ID:
        # Remove the "ORDER ID" column
        order_df.drop(columns="ORDER ID", inplace=True)

        # Sort the items by item number
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        # Append a "GRAND TOTAL" row
        total_price_df = pd.DataFrame({"ITEM PRICE":['GRAND TOTAL'],"TOTAL PRICE":[order_df[['TOTAL PRICE']].sum()]})
        
        order_df = pd.concat([order_df,total_price_df])

        # Determine the file name and full path of the Excel sheet
        customer_name = order_df['CUSTOMER NAME'].values[0]
        customer_name = re.sub(r'\W', '', customer_name)
        order_file_name = f'Order{order_id}_{customer_name}.xlsx'
        order_file_path = os.path.join(orders_dir, order_file_name)



        # Export the data to an Excel sheet
        order_df.to_excel(order_file_path, sheet_name=f"Order {order_id}")

        # TODO: Format the Excel sheet
    
    pass

if __name__ == '__main__':
    main()