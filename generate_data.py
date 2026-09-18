import pandas as pd
import numpy as np

np.random.seed(42)

n = 1000

dates = pd.date_range(
    start="2026-01-01",
    end="2026-09-18",
    periods=n
)

departments = [
    "Mobile",
    "TV",
    "Home Appliance",
    "IT"
]

products = [
    "Galaxy A",
    "Galaxy S",
    "Neo TV",
    "QLED TV",
    "Refrigerator",
    "Washing Machine",
    "Monitor",
    "Laptop"
]

df = pd.DataFrame({
    "Date": dates,
    "Department": np.random.choice(departments, n),
    "Product": np.random.choice(products, n),
    "Units": np.random.randint(1, 50, n),
    "Sales": np.random.randint(500, 30000, n),
    "Cost": np.random.randint(300, 22000, n),
    "Inventory": np.random.randint(10, 300, n)
})

df.to_excel("data/sales_data.xlsx", index=False)

print(df.head())
print("\nData shape:", df.shape)