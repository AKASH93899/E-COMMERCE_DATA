import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load the dataset
data = pd.read_excel("C:\\Users\\Akash Sharma\\OneDrive\\Desktop\\E-COMMERCE DATA.xlsx")
print(data.head(4))  # Preview the first 4 rows of the data

# Data cleaning: Replace spaces in the "Order Priority" column
data["Order Priority"] = data["Order Priority"].replace("Critical ","Critical")

# ------------------ Analysis and Visualization --------------------

# 1. Order Priority Distribution
plt.figure(figsize=(6,4))
sns.countplot(x="Order Priority", data=data)
plt.title("Count of Order Priority", fontsize=14)
plt.xlabel("Order Priority", fontsize=12)
plt.ylabel("Count", fontsize=12)
plt.tight_layout()
plt.savefig("count_of_order_priority.jpg")  # Save the figure
plt.show()

# 2. Ship Mode Distribution (Pie Chart)
ship_mode_counts = data["Ship Mode"].value_counts()
x = ship_mode_counts.index
y = ship_mode_counts.values
plt.figure(figsize=(6,4))
plt.pie(y, labels=x, startangle=55, autopct="%0.2f%%", wedgeprops={'edgecolor': 'black'})
plt.title("Distribution of Ship Modes", fontsize=14)
plt.legend(loc=2)
plt.tight_layout()
plt.show()

# 3. Ship Mode vs Product Category
plt.figure(figsize=(6,5))
sns.countplot(x="Ship Mode", data=data, hue="Product Category")
plt.title("Ship Mode Distribution by Product Category", fontsize=14)
plt.xlabel("Ship Mode", fontsize=12)
plt.ylabel("Count", fontsize=12)
plt.tight_layout()
plt.show()

# 4. Customer Segment Distribution
plt.figure(figsize=(6,5))
sns.countplot(x="Customer Segment", data=data)
plt.title("Distribution of Customer Segments", fontsize=14)
plt.xlabel("Customer Segment", fontsize=12)
plt.ylabel("Count", fontsize=12)
plt.tight_layout()
plt.show()

# 5. Top 5 Product Sub-Categories
category_counts = data['Product Sub-Category'].value_counts()
top_5_categories = category_counts.head(5)
colors = ["red", "blue", "pink", "black", "orange"]
plt.figure(figsize=(8,6))
top_5_categories.plot(kind='pie', colors=colors, autopct='%1.1f%%', startangle=90, wedgeprops={'edgecolor': 'black'})
plt.title('Top 5 Product Sub-Categories by Frequency', fontsize=15)
plt.tight_layout()
plt.show()

# 6. Orders by Year (Trend over Time)
data["Order Date"] = pd.to_datetime(data["Order Date"])  # Ensure datetime conversion
data["Order year"] = data["Order Date"].dt.year
plt.figure(figsize=(8, 6))
sns.countplot(x="Order year", data=data, palette='Blues_d')
plt.title('Count of Orders by Year', fontsize=16)
plt.xlabel('Year', fontsize=12)
plt.ylabel('Number of Orders', fontsize=12)
plt.xticks(rotation=0)
plt.tight_layout()
plt.show()

# 7. Total Profit by Product Category
plt.figure(figsize=(6,4))
sns.barplot(x="Product Category", y="Profit", data=data, estimator="sum", palette="Set2")
plt.xticks(rotation=45, ha='right')
plt.title("Total Profit by Product Category", fontsize=14)
plt.xlabel("Product Category", fontsize=12)
plt.ylabel("Total Profit", fontsize=12)
plt.tight_layout()
plt.show()

# 8. Top 5 States or Provinces
top_states = data["State or Province"].value_counts().head(5)
plt.figure(figsize=(8, 6))
sns.barplot(x=top_states.index, y=top_states.values, palette="viridis")
plt.title("Top 5 States or Provinces", fontsize=14)
plt.xlabel("State or Province", fontsize=12)
plt.ylabel("Count", fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# 9. Total Profit by Region
region_sum = data.groupby("Region")["Profit"].sum().reset_index()
plt.figure(figsize=(6, 5))
sns.barplot(x="Region", y="Profit", data=region_sum, palette="Set2")
plt.title("Total Profit by Region", fontsize=14)
plt.xlabel("Region", fontsize=12)
plt.ylabel("Total Profit", fontsize=12)
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# Additional Data Cleaning: Handle missing values (if necessary)
# Replace missing values in "Product Base Margin" with the mean value
data["Product Base Margin"] = data["Product Base Margin"].fillna(data["Product Base Margin"].mean())

# Show any missing data statistics (useful for further analysis)
print("Missing values per column:")
print(data.isnull().sum())

# Check for duplicate rows (if necessary)
dup = data.duplicated().sum()
print(f"Duplicate rows in the dataset: {dup}")
data.to_excel("C:\\Users\\Akash Sharma\\OneDrive\\Desktop\\E-COMMERCE_DATA.xlsx")