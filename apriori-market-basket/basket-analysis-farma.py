# -*- coding: utf-8 -*-
"""
Created on Mon Mar  2 16:46:56 2020

@author: Alex
"""
import pandas as pd
#from mlxtend.frequent_patterns import apriori
#from mlxtend.preprocessing import TransactionEncoder
#from mlxtend.frequent_patterns import association_rules
import numpy as np
import sys
from itertools import combinations, groupby
from collections import Counter
from IPython.display import display

def encode_units(x):
    if x <= 0:
        return 0
    if x >= 1:
        return 1

# Function that returns the size of an object in MB
def size(obj):
    return "{0:.2f} MB".format(sys.getsizeof(obj) / (1000 * 1000))

# Returns frequency counts for items and item pairs
def freq(iterable):
    if type(iterable) == pd.core.series.Series:
        return iterable.value_counts().rename("freq")
    else: 
        return pd.Series(Counter(iterable)).rename("freq")

    
# Returns number of unique orders
def order_count(order_item):
    return len(set(order_item.index))


# Returns generator that yields item pairs, one at a time
def get_item_pairs(order_item):
    order_item = order_item.reset_index().to_numpy()
    for order_id, order_object in groupby(order_item, lambda x: x[0]):
        item_list = [item[1] for item in order_object]
              
        for item_pair in combinations(item_list, 2):
            yield item_pair
            

# Returns frequency and support associated with item
def merge_item_stats(item_pairs, item_stats):
    return (item_pairs
                .merge(item_stats.rename(columns={'freq': 'freqA', 'support': 'supportA'}), left_on='item_A', right_index=True)
                .merge(item_stats.rename(columns={'freq': 'freqB', 'support': 'supportB'}), left_on='item_B', right_index=True))


# Returns name associated with item
def merge_item_name(rules, item_name):
    columns = ['itemA','itemB','freqAB','supportAB','freqA','supportA','freqB','supportB', 
               'confidenceAtoB','confidenceBtoA','lift']
    rules = (rules
                .merge(item_name.rename(columns={'item_name': 'itemA'}), left_on='item_A', right_on='item_id')
                .merge(item_name.rename(columns={'item_name': 'itemB'}), left_on='item_B', right_on='item_id'))
    return rules[columns]      

def association_rules(order_item, min_support):

    print("Starting order_item: {:22d}".format(len(order_item)))


    # Calculate item frequency and support
    item_stats             = freq(order_item).to_frame("freq")
    item_stats['support']  = item_stats['freq'] / order_count(order_item) * 100    


    # Filter from order_item items below min support 
    qualifying_items       = item_stats[item_stats['support'] >= min_support].index
    order_item             = order_item[order_item.isin(qualifying_items)]

    print("Items with support >= {}: {:15d}".format(min_support, len(qualifying_items)))
    print("Remaining order_item: {:21d}".format(len(order_item)))


    # Filter from order_item orders with less than 2 items
    order_size             = freq(order_item.index)
    qualifying_orders      = order_size[order_size >= 2].index
    order_item             = order_item[order_item.index.isin(qualifying_orders)]

    print("Remaining orders with 2+ items: {:11d}".format(len(qualifying_orders)))
    print("Remaining order_item: {:21d}".format(len(order_item)))


    # Recalculate item frequency and support
    item_stats             = freq(order_item).to_frame("freq")
    item_stats['support']  = item_stats['freq'] / order_count(order_item) * 100


    # Get item pairs generator
    item_pair_gen          = get_item_pairs(order_item)


    # Calculate item pair frequency and support
    item_pairs              = freq(item_pair_gen).to_frame("freqAB")
    item_pairs['supportAB'] = item_pairs['freqAB'] / len(qualifying_orders) * 100

    print("Item pairs: {:31d}".format(len(item_pairs)))


    # Filter from item_pairs those below min support
    item_pairs              = item_pairs[item_pairs['supportAB'] >= min_support]

    print("Item pairs with support >= {}: {:10d}\n".format(min_support, len(item_pairs)))


    # Create table of association rules and compute relevant metrics
    item_pairs = item_pairs.reset_index().rename(columns={'level_0': 'item_A', 'level_1': 'item_B'})
    item_pairs = merge_item_stats(item_pairs, item_stats)
    
    item_pairs['confidenceAtoB'] = item_pairs['supportAB'] / item_pairs['supportA']
    item_pairs['confidenceBtoA'] = item_pairs['supportAB'] / item_pairs['supportB']
    item_pairs['lift']           = item_pairs['supportAB'] / (item_pairs['supportA'] * item_pairs['supportB'])
    
    
    # Return association rules sorted by lift in descending order
    return item_pairs.sort_values('lift', ascending=False)

df = pd.read_excel('./Ventas/Ventas-2020.xls',header=0,converters={'Código':str})
df = df.append(pd.read_excel('./Ventas/Ventas-2019D.xls',header=0,converters={'Código':str}),ignore_index=True)
df = df.append(pd.read_excel('./Ventas/Ventas-2019C.xls',header=0,converters={'Código':str}),ignore_index=True)
df = df.append(pd.read_excel('./Ventas/Ventas-2019B.xls',header=0,converters={'Código':str}),ignore_index=True)
df = df.append(pd.read_excel('./Ventas/Ventas-2019A.xls',header=0,converters={'Código':str}),ignore_index=True)

df_orders = df.rename(columns = {'Código':'Code','Denominación':'product_name','Cantidad (Unidades)':'product_quantity'})

df_orders['order_id'] = (df_orders['Fecha']+'-'+df_orders['Hora']+'-'+df_orders['Vendedor'])

df_orders['product_name'] = df_orders['product_name'].str.strip()


cols = ['order_id','product_name','product_quantity']
df_orders = df_orders[cols]
df_orders = df_orders.dropna()
df_orders.sort_values(by=['order_id', 'product_name'], inplace = True)
print(df_orders.head(50))
# dropping ALL duplicte values 
df_orders.drop_duplicates(subset=['order_id', 'product_name'],keep = 'first', inplace = True) 

df_orders = df_orders.set_index('order_id')['product_name'].rename('item_id')
print(df_orders.head(50))
type(df_orders)

print('dimensions: {0};   size: {1};   unique_orders: {2};   unique_items: {3}'.format(df_orders.shape, size(df_orders), len(df_orders.index.unique()), len(df_orders.value_counts())))


rules = association_rules(df_orders, 0.01)

#basket = df_orders.groupby(['order_id','product_name'])['product_quantity'].sum().unstack().reset_index().fillna(0).set_index('order_id')
#basket_sets = basket.applymap(encode_units)

#frequent_itemsets = apriori(basket_sets, min_support=0.05, use_colnames=True)
#frequent_itemsets['length'] = frequent_itemsets['itemsets'].apply(lambda x: len(x))

#rules = association_rules(frequent_itemsets, metric="lift", min_threshold=1)
#print(rules.head())
