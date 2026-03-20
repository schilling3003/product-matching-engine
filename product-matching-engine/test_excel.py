import pandas as pd

# Create a simple test Excel file
data = {
    'Product Name': ['Apple iPhone 14', 'Samsung Galaxy S23', 'Google Pixel 7'],
    'Brand': ['Apple', 'Samsung', 'Google'],
    'Price': [999, 899, 699]
}

df = pd.DataFrame(data)
df.to_excel('test_products.xlsx', index=False)
print("Test Excel file created: test_products.xlsx")
