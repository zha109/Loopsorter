import pandas as pd
import os

data = {
    "OrderID": [1, 2, 3, 4, 5],
    "Wave": [1, 1, 2, 2, 3],
    "ReleaseTime": [0, 2, 5, 6, 10],
    "SKU": ["A", "B", "C", "D", "E"],
    "Quantity": [10, 5, 8, 12, 7],
    "PackingTime": [2, 3, 2, 4, 3],
    "LaneSpeed": [1.0, 1.2, 1.0, 1.1, 0.9]
}

df = pd.DataFrame(data)

file_path = os.path.join(os.getcwd(), "orders.xlsx")
df.to_excel(file_path, index=False)

print(f"File 'orders.xlsx' đã được tạo tại: {file_path}")
