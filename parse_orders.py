import argparse
import csv
from config import modify_list

orders = []
ORDER_ITEMS = 'מוצרים'
ITEMS_SOLD  = 'Items sold'
NET_REVENUE = 'N. הכנסה'

product_names = []

def parse_order_items(row):
    try:
        row[NET_REVENUE] = row[NET_REVENUE].replace("'", "")
        row[ITEMS_SOLD] = row[ITEMS_SOLD].replace("'", "")
        orig_items = row[ORDER_ITEMS]           
        items = orig_items
        row.pop(ORDER_ITEMS)
        if not orig_items:
            return
        #print(items)
        # Loop over modify list to change item names with comma into something else 
        for t in modify_list:
            items = items.replace(t[0], t[1])
        parsed_items = items.split(",")
        for item in parsed_items:
            #print(item)
            amount, name = item.split("×")
            # Loop over modify list to change back the item with comma in their names
            for t in modify_list:
                name = name.replace(t[1], t[0])
            name = name.strip()
            amount = amount.strip()
            if name not in product_names:
                product_names.append(name)
            #print(f"name {name} amount {amount}")
            row[name] = amount
    except Exception as e:
        print(orig_items)
        print(e)


def parse_orders(input_file_name, output_file_name):

    with open(input_file_name, 'r') as file:
        reader = csv.DictReader(file)

        for row in reader:
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


def main():
    parser = argparse.ArgumentParser(description='Parse Woocommerce orders')
    parser.add_argument('input_file', type=str, help='Input file - CSV you got from Woocommerce')
    parser.add_argument('output_file', type=str, help='Output file - name of output file to create')
    args = parser.parse_args()
    parse_orders(args.input_file, args.output_file)
    
if __name__ == '__main__':
    main()