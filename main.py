import pandas as pd
import re

def extract_symbol(details):
    """
    Extracts the symbol from the details text.
    This example regex looks for a word with 3 or more uppercase letters.
    Adjust the pattern as needed for your actual data.
    """
    match = re.search(r'\b([A-Z]{3,})\b', details)
    if match:
        return match.group(1)
    return ""

def extract_quantity(details):
    """
    Extracts the quantity from the details text.
    Looks for the words "Bought" or "Sold" followed by a number.
    The number may contain commas and/or a decimal point.
    """
    match = re.search(r'(Bought|Sold)\s+([\d,\.]+)', details)
    if match:
        quantity_str = match.group(2).replace(',', '')
        try:
            # Use float to handle both integer and decimal quantities.
            return float(quantity_str)
        except ValueError:
            return None
    return None

def extract_activity_type(details):
    """
    Determines the activity type by inspecting the details text.
    Returns one of the allowed types: BUY, SELL, DIVIDEND, INTEREST, DEPOSIT, WITHDRAWAL, FEE.
    """
    details_lower = details.lower()
    if 'bought' in details_lower:
        return 'BUY'
    elif 'sold' in details_lower:
        return 'SELL'
    elif 'dividend' in details_lower:
        return 'DIVIDEND'
    elif 'interest' in details_lower:
        return 'INTEREST'
    elif 'deposit' in details_lower:
        return 'DEPOSIT'
    elif 'withdrawal' in details_lower:
        return 'WITHDRAWAL'
    elif 'fee' in details_lower:
        return 'FEE'
    else:
        # If no known keyword is found, you might set this to a default or skip such rows.
        return 'UNKNOWN'

def calculate_unit_price(amount, quantity):
    """
    Calculates the unit price as Amount divided by Quantity.
    Returns None if quantity is zero or missing.
    """
    try:
        if quantity and quantity != 0:
            return amount / quantity
        else:
            return None
    except Exception:
        return None

def main():
    # Read the source CSV file
    try:
        x = input('write file name:')
        df = pd.read_csv(x)
    except Exception as e:
        print("Error reading the CSV file:", e)
        return

    # Convert the Date column to datetime and reformat it as YYYY-MM-DD
    try:
        df['Date'] = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
    except Exception as e:
        print("Error processing the Date column:", e)

    # Ensure that the Details column is processed as text
    df['Details'] = df['Details'].astype(str)

    # Extract Symbol, Quantity, and ActivityType from the Details column
    df['Symbol'] = df['Details'].apply(extract_symbol)
    df['Quantity'] = df['Details'].apply(extract_quantity)
    df['ActivityType'] = df['Details'].apply(extract_activity_type)

    # Calculate unitPrice = Amount / Quantity
    df['unitPrice'] = df.apply(lambda row: calculate_unit_price(row['Amount'], row['Quantity']), axis=1)

    # Set currency as a constant "GBP"
    df['currency'] = "GBP"

    # The final CSV should contain the following columns in order:
    # Date, Symbol, Quantity, ActivityType, unitPrice, currency, fee, Amount

    # Check for fee column name variations (it could be "Fee" or "fee")
    fee_col = 'Fee' if 'Fee' in df.columns else ('fee' if 'fee' in df.columns else None)
    if fee_col is None:
        print("Fee column not found in the CSV. Please verify the source file.")
        return

    final_columns = ['Date', 'Symbol', 'Quantity', 'ActivityType', 'unitPrice', 'currency', fee_col, 'Amount']
    final_df = df[final_columns]

    # Optionally, you can rename the fee column to lowercase "fee" for consistency
    final_df = final_df.rename(columns={fee_col: 'fee'})

    # Save the output to a new CSV file
    output_filename = 'output.csv'
    try:
        final_df.to_csv(output_filename, index=False)
        print(f"Output CSV file created successfully: {output_filename}")
    except Exception as e:
        print("Error writing the output CSV file:", e)

if __name__ == "__main__":
    main()
