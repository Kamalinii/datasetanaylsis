import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Define the file path
file_path = 'C:/Users/Admin/Downloads/sales_data.xlsx'

# Check if the file exists
if os.path.exists(file_path):
    # Load the dataset from the Excel file
    df = pd.read_excel(file_path)

    # Display the first few rows of the dataset
    print(df.head())

    # Filter sales records with sales amount greater than $1000
    filtered_df = df[df['Sales'] > 1000]
    print(filtered_df.head())

    # Sort the data by sales amount in descending order
    sorted_df = df.sort_values(by='Sales', ascending=False)
    print(sorted_df.head())

    # Group by region and calculate the total sales
    grouped_df = df.groupby('Region')['Sales'].sum().reset_index()
    print(grouped_df)

    # Calculate summary statistics for sales
    mean_sales = df['Sales'].mean()
    median_sales = df['Sales'].median()
    std_sales = df['Sales'].std()

    print(f"Mean Sales: {mean_sales}")
    print(f"Median Sales: {median_sales}")
    print(f"Standard Deviation of Sales: {std_sales}")

    # Set the style for the plots
    sns.set(style="whitegrid")

    # Plot the distribution of sales amounts
    plt.figure(figsize=(10, 6))
    sns.histplot(df['Sales'], bins=30, kde=True)
    plt.title('Distribution of Sales Amounts')
    plt.xlabel('Sales Amount')
    plt.ylabel('Frequency')
    plt.show()

    # Plot the relationship between sales and another numeric variable, such as 'Profit'
    plt.figure(figsize=(10, 6))
    sns.scatterplot(x='Sales', y='Profit', data=df)
    plt.title('Sales vs Profit')
    plt.xlabel('Sales Amount')
    plt.ylabel('Profit')
    plt.show()

    # Plot total sales by region
    plt.figure(figsize=(10, 6))
    sns.barplot(x='Region', y='Sales', data=grouped_df)
    plt.title('Total Sales by Region')
    plt.xlabel('Region')
    plt.ylabel('Total Sales')
    plt.show()
else:
    print(f"File not found: {file_path}")
