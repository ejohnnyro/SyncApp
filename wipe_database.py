import sqlite3

# Connect to the database
conn = sqlite3.connect('products.db')
cursor = conn.cursor()

# Execute SQL commands to delete all records
cursor.execute('DELETE FROM products')
cursor.execute('DELETE FROM vendor_stocks')

# Commit changes and close the connection
conn.commit()
conn.close()

print('Database wiped successfully.')