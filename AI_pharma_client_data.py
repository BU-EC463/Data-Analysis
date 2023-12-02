import pandas as pd
import matplotlib.pyplot as plt

'''
Purchases-Purchases Daily-Run Date-11/01/2023 - Date Range : 11/01/2022 - 10/31/2023
Columns:

ABC #	
Product Description	
NDC	
Account #	
Account Name	
Contract Abbrev Name	
Customer PO #	
Invoice #	
Invoice Date	
Order Qty	
Shipped Qty	
Invoice Ext Amt	
Total Ext Cost	
Invoice Price	
DEA Class	
GCN	
GCN Seq #	
Generic Abbrev Desc	
Invoice Type Code	
Invoice Type Desc	
Invoice Month	
Invoice Year	
Unit Size Qty	
Unit Size Code	
Unit Strength Qty	
Unit Strength Code	
Primary Ingredient HIC4 Code	
Primary Ingredient HIC4 Desc	
Route Desc	
FDB Package Size Qty	
FDB AWP Wholesale Factor	
NIOSH Code	
Abbreviated Desc	
Wholesale Cost	
Supplier Name	
AWP

NOTES: Most ordered products per active ingredient
'''

file_path = '/Users/bora/Downloads/AI_pharma_client_data.xlsx'


df = pd.read_excel(file_path, sheet_name='Results')


unique_count = df['ABC #'].nunique()
unique_count2 = df['Product Description'].nunique()
num_columns = df.shape[1]
product_orders = df.groupby('Product Description')['Order Qty'].sum()
product_cost_orders = df.groupby('Product Description')['Total Ext Cost'].sum()
product_cost_orders_average = df.groupby('Product Description')['Total Ext Cost'].mean()
primary_ingredient_orders = df.groupby('Primary Ingredient HIC4 Desc')['Order Qty'].sum()
most_ordered_product = product_orders.idxmax()
max_quantity_ordered = product_orders.max()
least_ordered_product = product_orders.idxmin()
min_quantity_ordered = product_orders.min()
most_ordered_ingredient = primary_ingredient_orders.idxmax()
max_quantity_ingredient = primary_ingredient_orders.max()
least_ordered_ingredient = primary_ingredient_orders.idxmin()
min_quantity_ingredient = primary_ingredient_orders.min()

top_5_expensive = product_cost_orders_average.nlargest(5)

df['Invoice Date'] = pd.to_datetime(df['Invoice Date'])
product_order_info = df.groupby('Product Description')['Invoice Date'].agg(['nunique', 'max', 'min'])
product_order_info['TimeRange'] = (product_order_info['max'] - product_order_info['min']).dt.days
product_order_info_filtered = product_order_info[product_order_info['nunique'] > 1]
average_time_to_order = product_order_info_filtered['TimeRange'].mean()

print(f'Number of unique items (ABC #): {unique_count}')
print(f'Number of unique items (Product Description): {unique_count2}')
print(f'Number of dimensions: {num_columns}')
print(f'The most ordered product is {most_ordered_product} with a total of {max_quantity_ordered} units ordered.')
print(f'The least ordered product is {least_ordered_product} with a total of {min_quantity_ordered} units ordered.')
print(f'The most ordered ingredient is {most_ordered_ingredient} with a total of {max_quantity_ingredient} units ordered.')
print(f'The least ordered ingredient is {least_ordered_ingredient} with a total of {min_quantity_ingredient} units ordered.')
print(f'The most spent on a drug is ${product_cost_orders.max()} which is {product_cost_orders.idxmax()}.')
print(f'The least spent on a drug is ${product_cost_orders.min()} which is {product_cost_orders.idxmin()}.')
print(f'Top 5 most expensive drugs (average cost): {top_5_expensive}')
print(f'The least spent on a drug (average cost) is ${product_cost_orders_average.min()} which is {product_cost_orders_average.idxmin()}.')
print(f'The average time it takes to order each product (excluding products with one order date) is {average_time_to_order:.2f} days.')
print('Quickest moving drugs are', product_order_info_filtered['TimeRange'].nsmallest(5))

#Drugs wrt their quantities
'''
plt.figure(figsize=(20, 10))
product_orders.plot(kind='bar')
plt.xlabel('Product')
plt.ylabel('Quantity')
plt.title('Product Quantities')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()
'''

#Drugs wrt their average cost
'''
plt.figure(figsize=(20, 10))
product_cost_orders_average.plot(kind='bar')
plt.xlabel('Product')
plt.ylabel('Average Cost')
plt.title('Product Average Cost')
plt.xticks(rotation=90)
plt.tight_layout()
plt.show()
'''