import argparse
import csv
from config import modify_list
import xlsxwriter

orders = []
ORDER_ITEMS = 'מוצרים'
ITEMS_SOLD  = 'Items sold'
NET_REVENUE = 'הכנסה נטו (מפורמט)'
DATE = 'תאריך'
SOLIDERS = 'הנחת חיילים'
REGULAR = 'הזמנה רגילה'
OUT = 'פלט'
SUMMARY = 'סיכום'


product_names = []

def parse_order_items(row):
    try:
        #print(row)
        row[NET_REVENUE] = row[NET_REVENUE].replace("'", "")
        row[ITEMS_SOLD] = row[ITEMS_SOLD].replace("'", "")

        orig_items = row[ORDER_ITEMS]           
        items = orig_items
        row.pop(ORDER_ITEMS)
        if not orig_items:
            print(f"No items in {row}")
            return
        #print(items)
        # Loop over modify list to change item names with comma into something else 
        for t in modify_list:
            items = items.replace(t[0], t[1])
        parsed_items = items.split(",")
        for item in parsed_items:
            if "×" not in item:
                print(f"Warning: couldn't find amount in item: {item}")
                for i, a in enumerate(items):
                    print(f"{i}: {a}")
                continue
            amount, name = item.split("×")
            # Loop over modify list to change back the item with comma in their names
            for t in modify_list:
                name = name.replace(t[1], t[0])
            name = name.strip()
            amount = amount.strip()
            if name not in product_names:
                product_names.append(name)
            #print(f"name {name} amount {amount}")
            if name == SOLIDERS:
                amount = 1
            row[name] = int(amount)
    except Exception as e:
        #print(row)
        print(f"Exception in row {row}")
        print(e)


def generate_summary(summary_file_name, orders):
    orders_summary = {}
    summary_products = []
    for product in product_names:
        if product != SOLIDERS:
            summary_products.append(product)
            orders_summary[product] = { REGULAR: 0, SOLIDERS: 0}

    for order in orders:
        key = REGULAR
        if SOLIDERS in order.keys():
            key = SOLIDERS
        for product in summary_products:
            if product in order.keys():
                orders_summary[product][key] += order[product]

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(summary_file_name)
    worksheet = workbook.add_worksheet()
    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    # Write the header
    worksheet.write(row, col,     ORDER_ITEMS)
    worksheet.write(row, col + 1, REGULAR)
    worksheet.write(row, col + 2, SOLIDERS)
    row += 1
    # Iterate over the data and write it out row by row.
    for product in summary_products:
        worksheet.write(row, col,     product)
        worksheet.write(row, col + 1, orders_summary[product][REGULAR])
        worksheet.write(row, col + 2, orders_summary[product][SOLIDERS])
        row += 1

    workbook.close()

def parse_orders(input_file_name, start_date):

    base_file_name = input_file_name.split(".")[0]
    output_file_name = base_file_name + "_" + OUT + ".csv"
    summary_file_name = base_file_name + "_" + SUMMARY + ".csv"

    with open(input_file_name, 'r', encoding='utf-8') as file:
        reader = csv.DictReader(file)

        for row in reader:
            # DATE is the first cell, but as there is a strange unicode issue with it,
            # just take the first key with this simple loop over the row keys
            # date = row[DATE]
            for k in row.keys():
                d = k
                break
            date = row[d]
            if start_date and start_date < date:
                #this will also populate the global product_names
                parse_order_items(row)
                orders.append(row)

        out_columns = reader.fieldnames
        out_columns.remove(ORDER_ITEMS)
        product_names.sort()
        out_columns += product_names
        #print(out_columns)
        
        
    with open(output_file_name, 'w', encoding='utf-8-sig', newline='') as output_file:
        writer = csv.DictWriter(output_file, fieldnames=out_columns, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(orders)

    generate_summary(summary_file_name, orders)

def main():
    parser = argparse.ArgumentParser(description='Parse Woocommerce orders')
    parser.add_argument('input_file', type=str, help='Input file - CSV you got from Woocommerce')
    parser.add_argument('--start_date', type=str, default="2023-01-01 00:00:00", help='Start processing from this date')
    args = parser.parse_args()
    parse_orders(args.input_file, args.start_date)
    
if __name__ == '__main__':
    main()