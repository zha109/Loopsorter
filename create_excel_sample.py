import pandas as pd

data = [
    {"OrderID": 1, "Wave": 1, "ReleaseTime": 0, "SKU": "SKU1", "Quantity": 10, "PackingTime": 2, "ProcessingTime": 5, "Lane": 1},
    {"OrderID": 2, "Wave": 1, "ReleaseTime": 1, "SKU": "SKU2", "Quantity": 15, "PackingTime": 3, "ProcessingTime": 6, "Lane": 1},
    {"OrderID": 3, "Wave": 2, "ReleaseTime": 3, "SKU": "SKU1", "Quantity": 5, "PackingTime": 2, "ProcessingTime": 4, "Lane": 2},
    {"OrderID": 4, "Wave": 2, "ReleaseTime": 5, "SKU": "SKU3", "Quantity": 8, "PackingTime": 3, "ProcessingTime": 5, "Lane": 2},
]

df = pd.DataFrame(data)
df.to_excel("orders.xlsx", index=False)
print("Sample orders.xlsx created!")
