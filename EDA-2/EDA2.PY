import pandas as pd
import logging
import matplotlib.pyplot as plt
import seaborn as sns

class OrderProcessing:
    def __init__(self, file_path):
        self.file_path = file_path
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
        self.handler = logging.FileHandler('order_processing.log')
        self.handler.setLevel(logging.ERROR)
        self.logger.addHandler(self.handler)

    def load_data(self):
        try:
            df = pd.read_csv(self.file_path)
            return df
        except Exception as e:
            self.logger.exception("Error occurred while loading data")
            return None

    def handle_missing_billing_address(self, df):
        try:
            df['Payment Type'] = df['Payment Method'].str.split('₹').str[0].str.strip()
            missing_billing = df[df['Billing Street Address'].isnull()]
            for index, row in missing_billing.iterrows():
                if pd.notnull(row['Payment Type']) and 'Offline' in row['Payment Type']:
                    df.at[index, 'Billing Name'] = row['Shipping Name']
                    df.at[index, 'Billing Country'] = row['Shipping Country']
                    df.at[index, 'Billing Street Address'] = row['Shipping Street Address']
                    df.at[index, 'Billing Street Address 2'] = row['Shipping Street Address 2']
                    df.at[index, 'Billing City'] = row['Shipping City']
                    df.at[index, 'Billing State'] = row['Shipping State']
                    df.at[index, 'Billing Zip'] = row['Shipping Zip']
            return df
        except Exception as e:
            self.logger.exception("Error occurred while handling missing billing address")
            return None

    def save_to_csv(self, df, file_name):
        try:
            df.to_csv(file_name, index=False)
            return True
        except Exception as e:
            self.logger.exception("Error occurred while saving data to CSV")
            return False

    def plot_top_items(self, df):
        try:
            item_quantities = df.groupby('LineItem Name')['LineItem Qty'].sum()
            top_20_items = item_quantities.sort_values(ascending=False).head(20)
            plt.figure(figsize=(10, 6))
            top_20_items.plot(kind='bar', color='skyblue')
            plt.title('Top 20 Highest Quantity Items')
            plt.xlabel('Item Name')
            plt.ylabel('Total Quantity')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            plt.show()
        except Exception as e:
            self.logger.exception("Error occurred while plotting top items")

    def plot_boxplot(self, df):
        try:
            plt.figure(figsize=(8, 6))
            sns.boxplot(data=df, y='LineItem Qty', color='skyblue')
            plt.title('Boxplot of LineItem Qty')
            plt.ylabel('Quantity')
            plt.show()
        except Exception as e:
            self.logger.exception("Error occurred while plotting boxplot")

    def handle_unusually_large_orders(self, df):
        try:
            outlier_orders = df[df['LineItem Qty'] > 5]
            plt.figure(figsize=(8, 6))
            sns.histplot(df['LineItem Qty'], bins=20, kde=True, color='skyblue', alpha=0.7)
            plt.title('Histogram of LineItem Qty')
            plt.xlabel('Quantity')
            plt.ylabel('Frequency')
            plt.axvline(x=5, color='red', linestyle='--', label='Threshold (5 items)')
            plt.legend()
            plt.show()
            return outlier_orders
        except Exception as e:
            self.logger.exception("Error occurred while handling unusually large orders")
            return None

    def handle_multiple_payment_addresses(self, df):
        try:
            df['Payment Type'] = df['Payment Method'].str.split('₹').str[0].str.strip()
            multiple_payment_addresses = df.groupby('Shipping Street Address').apply(lambda x: x['Payment Type'].nunique() > 1)
            multiple_payment_addresses = multiple_payment_addresses[multiple_payment_addresses].index
            rows_with_multiple_payment_types = df[df['Shipping Street Address'].isin(multiple_payment_addresses)]
            return rows_with_multiple_payment_types
        except Exception as e:
            self.logger.exception("Error occurred while handling multiple payment addresses")
            return None

    def handle_international_orders(self, df):
        try:
            international_orders = df[df['Shipping Country'] != 'IND']
            return international_orders
        except Exception as e:
            self.logger.exception("Error occurred while handling international orders")
            return None

file_path = 'orders_2020_2021_DataSet_Updated'
order_processor = OrderProcessing(file_path)
df = order_processor.load_data()
if df is not None:
    df = order_processor.handle_missing_billing_address(df)
    if df is not None:
        order_processor.save_to_csv(df, 'missing_billing_address.csv')
        top_items_df = order_processor.plot_top_items(df)
        if top_items_df is not None:
            order_processor.plot_boxplot(df)
            unusually_large_orders_df = order_processor.handle_unusually_large_orders(df)
            if unusually_large_orders_df is not None:
                order_processor.save_to_csv(unusually_large_orders_df, 'unusually_large_orders.csv')
                multiple_payment_addresses_df = order_processor.handle_multiple_payment_addresses(df)
                if multiple_payment_addresses_df is not None:
                    order_processor.save_to_csv(multiple_payment_addresses_df, 'multiple_payment_addresses.csv')
                    international_orders_df = order_processor.handle_international_orders(df)
                    if international_orders_df is not None:
                        order_processor.save_to_csv(international_orders_df, 'international_orders.csv')
