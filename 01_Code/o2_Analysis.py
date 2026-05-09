import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

df = pd.read_excel(r"C:\Users\INTEL\Desktop\Work\Amazon_Scraper\02_Data\Laptop_Data.xlsx")
print(df.head(10))

print("\n",df.describe())

# data overview
tnop = len(df) # tnop = total number of product
min_price = int(df["PRICE"].min())
max_price = int(df["PRICE"].max())
avg_price = int(df["PRICE"].mean())

# baxplot for statics
plt.figure(figsize=(12,6))
sns.boxplot(x="PRICE", data=df, 
    color="skyblue",
    linewidth=2)
plt.title("Laptop Price Boxplot", fontsize= 22)
plt.xlabel("Price (₹)", fontsize=12)
plt.xticks(np.arange(0, df["PRICE"].max()+20000, 20000), rotation=45)
plt.grid(axis="x", linestyle="--", alpha=0.5)
plt.show()

# Histogram graph for laptop price distribution
plt.figure(figsize=(12,6))
sns.histplot(df["PRICE"], bins=30, kde = True)
plt.xticks(np.arange(0, df["PRICE"].max()+20000, 20000), rotation=45)
plt.title("Laptop Price Distribution", fontsize = 22)
plt.xlabel("Price (₹)", fontsize=12)
plt.show()

# printing 10 cheapest laptop
top10_cheapest = df.sort_values("PRICE").head(10).reset_index(drop=True)
print("\nTop 10 Cheapest - \n", top10_cheapest)


# printing top 10 Expensive laptop
top10_expensive = df.sort_values(by="PRICE", ascending=False).head(10).reset_index(drop=True)
print("\nTop 10 Expensive - \n", top10_expensive)

# Average price by brand and saving it into original file in new sheet   
avg_price_brand = df.groupby("BRAND")["PRICE"].mean().round(2).sort_values().reset_index()
avg_price_brand.columns = ["Brand Name", "Average Price"]

with pd.ExcelWriter(r"C:\Users\INTEL\Desktop\Work\Amazon_Scraper\02_Data\Laptop_Data.xlsx", mode = "a", engine = "openpyxl", if_sheet_exists="replace")as writer:
    avg_price_brand.to_excel(writer, sheet_name="Average Price by Company", index = False)

# Bar chart for average price by company in excel sheets
wb = load_workbook(r"C:\Users\INTEL\Desktop\Work\Amazon_Scraper\02_Data\Laptop_Data.xlsx")
ws = wb["Average Price by Company"]

data_ws = Reference(ws, min_col=2, min_row=1, max_row=16) 
categories = Reference(ws, min_col=1, min_row=2, max_row=16)

chart = BarChart()
chart.add_data(data_ws, titles_from_data=True)
chart.set_categories(categories)
chart.style = 10
chart.title = "Revenue By Product"
ws.add_chart(chart, "E2")
wb.save(r"C:\Users\INTEL\Desktop\Work\Amazon_Scraper\02_Data\Laptop_Data.xlsx")

# brand barchart
plt.figure(figsize=(12,6))
sns.histplot(df["BRAND"], kde = True)
plt.title("Brand Distribution", fontsize = 22)
plt.xlabel("Brand", fontsize=12)
plt.show()