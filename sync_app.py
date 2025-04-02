import sys
import os
import tkinter as tk
import threading
from tkinter import ttk, messagebox
import requests
from dotenv import load_dotenv, set_key
from functools import partial
from database import DatabaseManager
from tqdm import tqdm
import json
import webbrowser
from database import DatabaseManager, Product
import queue
import time

class SettingsDialog:
    def __init__(self, app):
        self.dialog = tk.Toplevel(app.root)
        self.dialog.title("Settings")
        self.dialog.geometry("650x500")
        self.dialog.transient(app.root)
        self.dialog.grab_set()
        self.app = app
        
        # Configure styles
        style = ttk.Style()
        style.configure("Settings.TLabel", font=("Helvetica", 12))
        style.configure("SettingsTitle.TLabel", font=("Helvetica", 14, "bold"))
        style.configure("SettingsHeader.TLabel", font=("Helvetica", 11, "bold"))
        
        # Main container
        main_container = ttk.Frame(self.dialog, padding="20 15 20 15")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        ttk.Label(title_frame, text="Application Settings", style="SettingsTitle.TLabel").pack()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # General Settings tab
        self.general_frame = ttk.Frame(self.notebook, padding="20 15 20 15")
        self.notebook.add(self.general_frame, text="General")
        
        # Section title for general frame
        ttk.Label(self.general_frame, text="General Settings", style="SettingsHeader.TLabel").pack(pady=(0, 15))
        ttk.Label(self.general_frame, text="Additional general settings will be available soon.", style="Settings.TLabel").pack(pady=10)
        
        # WooCommerce API Settings tab
        self.woo_frame = ttk.Frame(self.notebook, padding="20 15 20 15")
        self.notebook.add(self.woo_frame, text="WooCommerce API")
        
        # Section title
        ttk.Label(self.woo_frame, text="API Configuration", style="SettingsHeader.TLabel").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))
        
        # Grid layout for WooCommerce API configuration
        ttk.Label(self.woo_frame, text="Store URL:", style="Settings.TLabel").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=8)
        self.url_input = ttk.Entry(self.woo_frame, width=50)
        self.url_input.insert(0, self.app.url_input.get())
        self.url_input.grid(row=1, column=1, sticky=tk.EW)
        
        ttk.Label(self.woo_frame, text="API Key:", style="Settings.TLabel").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=8)
        self.key_input = ttk.Entry(self.woo_frame, width=50)
        self.key_input.insert(0, self.app.key_input.get())
        self.key_input.grid(row=2, column=1, sticky=tk.EW)
        
        ttk.Label(self.woo_frame, text="API Secret:", style="Settings.TLabel").grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=8)
        self.secret_input = ttk.Entry(self.woo_frame, width=50, show="*")
        self.secret_input.insert(0, self.app.secret_input.get())
        self.secret_input.grid(row=3, column=1, sticky=tk.EW)
        
        # Test connection button with improved styling
        button_frame = ttk.Frame(self.woo_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        self.test_button = ttk.Button(button_frame, text="Test Connection",
                                    command=lambda: self.app.test_connection(self.url_input.get(),
                                                                          self.key_input.get(),
                                                                          self.secret_input.get(),
                                                                          self))
        self.test_button.pack()
        
        # Create status label with better positioning
        self.status_label = ttk.Label(self.woo_frame, text="", style="Status.TLabel")
        self.status_label.grid(row=5, column=0, columnspan=2, pady=(0, 10))
        
        # Configure grid weights
        self.woo_frame.grid_columnconfigure(1, weight=1)
        
        # Vendor APIs tab with improved layout
        self.vendor_frame = ttk.Frame(self.notebook, padding="20 15 20 15")
        self.notebook.add(self.vendor_frame, text="Vendor APIs")
        
        # Section title for vendor frame
        ttk.Label(self.vendor_frame, text="Vendor API Integration", style="SettingsHeader.TLabel").pack(pady=(0, 15))
        ttk.Label(self.vendor_frame, text="Additional vendor API integration will be available soon.", style="Settings.TLabel").pack(pady=10)
        
        # Buttons frame with improved styling
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=20, pady=15)
        
        ttk.Button(button_frame, text="Save", command=self.save_settings, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side=tk.RIGHT, padx=5)
    
    def save_settings(self):
        # Update main window's stored credentials
        self.app.url_input.set(self.url_input.get())
        self.app.key_input.set(self.key_input.get())
        self.app.secret_input.set(self.secret_input.get())
        
        # Save to .env file
        set_key('.env', 'WOO_API_URL', self.url_input.get())
        set_key('.env', 'WOO_API_KEY', self.key_input.get())
        set_key('.env', 'WOO_API_SECRET', self.secret_input.get())
        
        # Reload environment variables
        load_dotenv(override=True)
        
        self.dialog.destroy()

class SyncApp:
    def sort_treeview(self, col):
        # Get all items in the treeview
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]
        
        if not items:
            return
        
        # Determine if the column should be sorted numerically
        numeric_columns = {'Regular Price', 'Sale Price', 'Stock'}
        is_numeric = col in numeric_columns
        
        # Sort items based on column type
        if is_numeric:
            # For numeric columns, handle 'N/A' and None values
            def numeric_sort_key(x):
                val = x[0]
                if val in ('N/A', 'None', ''):
                    return float('-inf')
                try:
                    return float(val)
                except ValueError:
                    return float('-inf')
            
            items.sort(key=numeric_sort_key, reverse=getattr(self, 'sort_reverse', False))
        else:
            # For text columns, use case-insensitive string comparison
            items.sort(key=lambda x: x[0].lower() if isinstance(x[0], str) and x[0] not in ('N/A', 'None') else '',
            reverse=getattr(self, 'sort_reverse', False))
        
        # Rearrange items in sorted positions
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)
            # Maintain link style for ID column
            if col != 'ID' and self.tree.set(item, 'ID') != 'N/A':
                self.tree.item(item, tags=('link',))
        
        # Reverse sort next time
        self.sort_reverse = not getattr(self, 'sort_reverse', False)
        
        # Update column headers to show sort direction
        for column in self.tree['columns']:
            self.tree.heading(column, text=column)
        self.tree.heading(col, text=f"{col} {'↓' if self.sort_reverse else '↑'}")

    def update_product_list(self, filtered_products=None):
        print("\nUpdating product list in UI...")
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get products from database with pagination
        page = getattr(self, 'current_page', 0)
        per_page = 50
        offset = page * per_page
        
        if filtered_products is not None:
            products_to_display = filtered_products[offset:offset + per_page]
            total_products = len(filtered_products)
        else:
            search_term = self.search_var.get() if hasattr(self, 'search_var') else None
            products_to_display = self.db.search_products(search_term=search_term, limit=per_page, offset=offset)
            total_products = self.db.get_total_products()
        
        print(f"Number of products to display: {len(products_to_display)}")
        
        # Add items
        products_added = 0
        for product in products_to_display:
            try:
                # Format price with currency if available
                regular_price = product.regular_price if isinstance(product, Product) else product.get('regular_price')
                sale_price = product.sale_price if isinstance(product, Product) else product.get('sale_price')
                
                if regular_price is not None:
                    try:
                        # Extract numeric value from SQLAlchemy Column or direct value
                        if hasattr(regular_price, 'value'):
                            price_value = float(str(regular_price.value))
                        else:
                            price_value = float(str(regular_price))
                        
                        if self.show_tva_var.get():
                            price_value = price_value * 1.19
                        regular_price = f"{price_value:.2f}"
                    except (ValueError, TypeError, AttributeError):
                        regular_price = 'N/A'
                else:
                    regular_price = 'N/A'
                
                if sale_price is not None:
                    try:
                        # Extract numeric value from SQLAlchemy Column or direct value
                        if hasattr(sale_price, 'value'):
                            price_value = float(str(sale_price.value))
                        else:
                            price_value = float(str(sale_price))
                            
                        if self.show_tva_var.get():
                            price_value = price_value * 1.19
                        sale_price = f"{price_value:.2f}"
                    except (ValueError, TypeError, AttributeError):
                        sale_price = 'N/A'
                else:
                    sale_price = 'N/A'
                
                woo_id = product.woo_id if isinstance(product, Product) else product.get('id', 'N/A')
                name = product.name if isinstance(product, Product) else product.get('name', 'N/A')
                sku = product.sku if isinstance(product, Product) else product.get('sku', 'N/A')
                stock = product.stock_quantity if isinstance(product, Product) else product.get('stock_quantity', 'N/A')
                
                # Format last_synced timestamp
                last_synced = product.last_synced if isinstance(product, Product) else product.get('last_synced', 'N/A')
                if last_synced != 'N/A':
                    last_synced = last_synced.strftime('%Y-%m-%d %H:%M')
                
                values = (woo_id, name, sku, regular_price, sale_price, stock, last_synced)
                print(f"Adding product: {values}")
                
                item = self.tree.insert("", tk.END, values=values)
                if str(woo_id) != 'N/A':
                    self.tree.item(item, tags=('link',))
                products_added += 1
            except Exception as e:
                print(f"Error adding product to list: {str(e)}")
                print(f"Problematic product data: {product}")
                continue
        
        # Update pagination controls
        total_pages = (total_products + per_page - 1) // per_page
        
        # Update button states
        self.prev_page_btn.config(state=tk.NORMAL if page > 0 else tk.DISABLED)
        self.next_page_btn.config(state=tk.NORMAL if page < total_pages - 1 else tk.DISABLED)
        
        # Update page counter
        self.page_counter_label.config(text=f"Page {page + 1} of {total_pages}")
        
        # Update jump buttons
        jumps = [-50, -25, -10, 10, 25, 50]
        for jump in jumps:
            target_page = page + jump
            self.jump_buttons[jump].config(state=tk.NORMAL if 0 <= target_page < total_pages else tk.DISABLED)
        
        print(f"Successfully added {products_added} products to the UI")

    def edit_product(self, product_id):
        messagebox.showinfo("Edit Product", f"Editing product {product_id}")

    def sync_product(self, woo_id):
        try:
            # Store the selected item before updating
            selected_items = self.tree.selection()
            
            # Load API credentials
            load_dotenv(override=True)
            url = os.getenv('WOO_API_URL', '').rstrip('/')
            key = os.getenv('WOO_API_KEY', '')
            secret = os.getenv('WOO_API_SECRET', '')
            
            # Fetch product data from WooCommerce
            response = requests.get(
                f"{url}/wp-json/wc/v3/products/{woo_id}",
                auth=(key, secret)
            )
            
            if response.status_code == 200:
                # Update local database with fetched data
                product_data = response.json()
                self.db.add_or_update_product(product_data)
                self.status_label.config(text=f"Product {woo_id} synced from WooCommerce successfully", foreground="green")
                
                # Refresh the product list
                self.update_product_list()
                
                # Restore the selection
                if selected_items:
                    for item in self.tree.get_children():
                        if self.tree.item(item)['values'][0] == woo_id:
                            self.tree.selection_set(item)
                            self.tree.see(item)
                            break
            else:
                self.status_label.config(text=f"Failed to fetch product {woo_id}: {response.status_code}", foreground="red")
        except Exception as e:
            self.status_label.config(text=f"Error syncing product: {str(e)}", foreground="red")
            print(f"Error syncing product: {str(e)}")

    def sync_to_woocommerce(self, woo_id):
        try:
            # Store the selected item before updating
            selected_items = self.tree.selection()
            
            # Get the product data from the database
            product = self.db.get_product_by_id(woo_id)
            if not product:
                self.status_label.config(text=f"Product {woo_id} not found in database", foreground="red")
                return
            
            # Load API credentials
            load_dotenv(override=True)
            url = os.getenv('WOO_API_URL', '').rstrip('/')
            key = os.getenv('WOO_API_KEY', '')
            secret = os.getenv('WOO_API_SECRET', '')
            
            # Prepare the data to update
            update_data = {
                'regular_price': str(product.regular_price) if product.regular_price is not None else '',
                'sale_price': str(product.sale_price) if product.sale_price is not None else '',
                'stock_quantity': product.stock_quantity
            }
            
            # Update product in WooCommerce
            response = requests.put(
                f"{url}/wp-json/wc/v3/products/{woo_id}",
                auth=(key, secret),
                json=update_data
            )
            
            if response.status_code == 200:
                self.status_label.config(text=f"Product {woo_id} updated in WooCommerce successfully", foreground="green")
                # Update last_synced timestamp in database
                self.db.update_product_sync_time(woo_id)
                # Refresh the product list
                self.update_product_list()
                # Restore the selection
                if selected_items:
                    for item in self.tree.get_children():
                        if self.tree.item(item)['values'][0] == woo_id:
                            self.tree.selection_set(item)
                            self.tree.see(item)
                            break
            else:
                self.status_label.config(text=f"Failed to update product {woo_id}: {response.status_code}", foreground="red")
        except Exception as e:
            self.status_label.config(text=f"Error updating product: {str(e)}", foreground="red")
            print(f"Error syncing product: {str(e)}")

    def show_context_menu(self, event):
        print(f"\nRight-click event detected at coordinates: ({event.x}, {event.y})")
        region = self.tree.identify('region', event.x, event.y)
        print(f"Clicked region: {region}")
        if region == 'cell':
            column = self.tree.identify_column(event.x)
            print(f"Clicked column: {column}")
            item = self.tree.identify('item', event.x, event.y)
            if item:
                values = self.tree.item(item)['values']
                product_id = values[0]  # WooCommerce ID
                
                try:
                    menu = tk.Menu(self.tree, tearoff=0)
                    
                    # Get the cell value for copying
                    col_idx = int(column.replace('#', '')) - 1
                    cell_value = str(values[col_idx]) if col_idx < len(values) else ''
                    
                    # Add Copy Value option for all cells
                    menu.add_command(label="Copy Value",
                                   command=lambda v=cell_value: self.copy_to_clipboard(v))
                    
                    if column == '#1' and product_id != 'N/A':  # ID column
                        # Construct product URL using base URL from environment
                        base_url = self.url_input.get().rstrip('/')
                        product_url = f"{base_url}/?p={product_id}"
                        menu.add_command(label="Open in Browser",
                                       command=lambda p=product_url: webbrowser.open(p))
                    
                    if column == '#2' and product_id != 'N/A':  # Name column
                        menu.add_command(label="Sync from WooCommerce",
                                       command=lambda pid=product_id: self.sync_product(pid))
                        # Add new Sync To WooCommerce option with red color
                        menu.add_command(label="Sync To WooCommerce",
                                       command=lambda pid=product_id: self.sync_to_woocommerce(pid),
                                       foreground="red")
                    
                    # Show menu at event coordinates if it has commands
                    if menu.index('end') is not None:
                        menu.tk_popup(event.x_root, event.y_root)
                except Exception as e:
                    print(f"Error showing context menu: {str(e)}")

    def on_motion(self, event):
        item = self.tree.identify('item', event.x, event.y)
        if item:
            column = self.tree.identify_column(event.x)
            if column == '#1':  # ID column
                self.root.config(cursor='hand2')
                return
        self.root.config(cursor='')
    
    def on_leave(self, event):
        self.root.config(cursor='')

    def on_double_click(self, event):
        # Get the item and column that was clicked
        region = self.tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
            
        column = self.tree.identify_column(event.x)
        item = self.tree.identify('item', event.x, event.y)
        
        if not item:
            return
            
        # Get column name
        col_id = int(column.replace('#', '')) - 1
        col_name = self.tree['columns'][col_id]
        
        # Only allow editing of price and stock columns
        if col_name not in ('Regular Price', 'Sale Price', 'Stock'):
            return
            
        # Get the current value and position
        current_value = self.tree.item(item)['values'][col_id]
        if current_value == 'N/A':
            current_value = ''
            
        # Get the bbox of the cell
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
            
        # Create an entry widget for editing
        entry = ttk.Entry(self.tree, width=len(str(current_value)) + 5)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        
        # Position the entry widget
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def on_entry_return(event):
            self.save_edit(item, col_name, entry.get())
            entry.destroy()
            
        def on_entry_escape(event):
            entry.destroy()
            
        def on_focus_out(event):
            entry.destroy()
            
        entry.bind('<Return>', on_entry_return)
        entry.bind('<Escape>', on_entry_escape)
        entry.bind('<FocusOut>', on_focus_out)
        entry.focus_set()
        
        # Store reference to prevent garbage collection
        self.edit_widget = entry

    def save_edit(self, item, column, value):
        try:
            # Get the product ID from the first column
            product_id = self.tree.item(item)['values'][0]
            if product_id == 'N/A':
                return

            # Validate and convert the value based on column type
            if column in ('Regular Price', 'Sale Price'):
                try:
                    value = float(value) if value else None
                except ValueError:
                    messagebox.showerror("Invalid Input", "Please enter a valid number for price.")
                    return
            elif column == 'Stock':
                try:
                    value = int(value) if value else 0
                except ValueError:
                    messagebox.showerror("Invalid Input", "Please enter a valid number for stock.")
                    return

            # Update the database
            column_map = {
                'Regular Price': 'regular_price',
                'Sale Price': 'sale_price',
                'Stock': 'stock_quantity'
            }
            db_column = column_map[column]
            self.db.update_product_field(product_id, db_column, value)

            # Update the tree view
            values = list(self.tree.item(item)['values'])
            col_id = self.tree['columns'].index(column)
            # Format display value based on column type
            if value is None:
                display_value = 'N/A'
            elif column in ('Regular Price', 'Sale Price'):
                display_value = f"{value:.2f}"
            else:
                display_value = str(value)
            values[col_id] = display_value
            self.tree.item(item, values=values)

            # Show success message
            self.status_label.config(text=f"{column} updated successfully", foreground="green")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update {column.lower()}: {str(e)}")

    def filter_products(self, *args):
        self.current_page = 0  # Reset to first page when searching
        self.update_product_list()
    
    def go_to_page(self, page):
        self.current_page = page
        self.update_product_list()
        
    def next_page(self):
        self.current_page += 1
        self.update_product_list()
        
    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_product_list()

    def copy_to_clipboard(self, value):
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.status_label.config(text=f"Value copied to clipboard", foreground="green")

    def load_tva_preference(self):
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
                return config.get('show_tva', True)
        except FileNotFoundError:
            return True
    
    def save_tva_preference(self):
        config = {'show_tva': self.show_tva_var.get()}
        with open('config.json', 'w') as f:
            json.dump(config, f)
    
    def toggle_tva(self):
        self.save_tva_preference()
        self.update_product_list()
    
    def __init__(self, root):
        self.root = root
        self.root.title("WooCommerce Product Sync")
        self.root.geometry("1024x768")
        self.root.state("zoomed")
        
        # Initialize database
        self.db = DatabaseManager()
        
        # Load API credentials from environment variables
        load_dotenv()
        self.url_input = tk.StringVar(value=os.getenv('WOO_API_URL', ''))
        self.key_input = tk.StringVar(value=os.getenv('WOO_API_KEY', ''))
        self.secret_input = tk.StringVar(value=os.getenv('WOO_API_SECRET', ''))
        
        # Initialize pagination
        self.current_page = 0
        
        # Configure style
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Helvetica", 16))
        style.configure("Status.TLabel", font=("Helvetica", 10))
        style.configure("Page.TLabel", font=("Helvetica", 10))
        style.configure("Title.TLabel", font=("Helvetica", 16, "bold"))
        style.configure("Status.TLabel", font=("Helvetica", 10))
        style.configure("Page.TLabel", font=("Helvetica", 10))
        
        # Create menu bar
        self.create_menu_bar()
        
        # Main container with padding
        main_container = ttk.Frame(root, padding="20 10 20 10")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_container, text="WooCommerce Product Management", style="Title.TLabel")
        title_label.pack(pady=(0, 20))
        
        # Status label
        self.status_label = ttk.Label(main_container, text="", style="Status.TLabel")
        self.status_label.pack(pady=5)
        
        # Search frame
        search_frame = ttk.Frame(main_container)
        search_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(search_frame, text="Search:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.filter_products)
        self.search_entry = ttk.Entry(search_frame, width=40, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT)
        
        # Fetch products button
        self.fetch_button = ttk.Button(search_frame, text="Fetch Products", command=self.fetch_products)
        self.fetch_button.pack(side=tk.RIGHT, padx=5)
        
        # Create table frame
        table_frame = ttk.Frame(main_container)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Products table with sorting
        self.tree = ttk.Treeview(table_frame, columns=("ID", "Name", "SKU", "Regular Price", "Sale Price", "Stock", "Last Synced"), show="headings", height=20)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Configure columns
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c))
            if col == "Name":
                self.tree.column(col, width=300, stretch=True)
            elif col == "ID":
                self.tree.column(col, width=50, stretch=False)
            elif col == "SKU":
                self.tree.column(col, width=120, stretch=False)
            elif col == "Regular Price":
                self.tree.column(col, width=100, stretch=False)
            elif col == "Sale Price":
                self.tree.column(col, width=100, stretch=False)
            elif col == "Stock":
                self.tree.column(col, width=100, stretch=False)
            elif col == "Last Synced":
                self.tree.column(col, width=150, stretch=False)
            else:
                self.tree.column(col, stretch=True)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Add pagination controls
        pagination_frame = ttk.Frame(main_container)
        pagination_frame.pack(fill=tk.X, pady=5)
        
        # Previous button
        self.prev_page_btn = ttk.Button(pagination_frame, text="Previous", command=self.prev_page)
        self.prev_page_btn.pack(side=tk.LEFT, padx=2)
        
        # Initialize jump buttons
        self.jump_buttons = {}
        jumps = [-50, -25, -10, 10, 25, 50]
        
        # Add jump buttons in order with page counter in the middle
        for jump in jumps:
            if jump == 10:
                # Page counter
                self.page_counter_label = ttk.Label(pagination_frame, text="Page 1 of 1", style="Page.TLabel")
                self.page_counter_label.pack(side=tk.LEFT, padx=5)
            
            btn = ttk.Button(pagination_frame, text=f"{jump:+d}", 
                            command=lambda j=jump: self.go_to_page(self.current_page + j))
            btn.pack(side=tk.LEFT, padx=2)
            self.jump_buttons[jump] = btn
        
        # Next button
        self.next_page_btn = ttk.Button(pagination_frame, text="Next", command=self.next_page)
        self.next_page_btn.pack(side=tk.LEFT, padx=2)
        
        # TVA checkbox
        self.show_tva_var = tk.BooleanVar(value=self.load_tva_preference())
        self.show_tva_checkbox = ttk.Checkbutton(pagination_frame, text="Prices with TVA",
                                                variable=self.show_tva_var,
                                                command=self.toggle_tva)
        self.show_tva_checkbox.pack(side=tk.LEFT, padx=(10, 2))
            
        # Configure link style and bind events for cursor change
        self.tree.tag_configure('link', foreground='blue')
        self.tree.bind('<Button-3>', self.show_context_menu)  # Right-click event
        self.tree.bind('<Motion>', self.on_motion)
        self.tree.bind('<Leave>', self.on_leave)
        self.tree.bind('<Double-Button-1>', self.on_double_click)  # Add double-click event
        self.tree.column("Name", width=300)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack the tree and scrollbar
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store products data
        self.sort_reverse = False
        
        # Display initial products
        self.update_product_list()
    
    def create_menu_bar(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Settings menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Settings", menu=settings_menu)
        settings_menu.add_command(label="API Configuration", command=self.open_settings)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def open_settings(self):
        SettingsDialog(self)
    
    def show_about(self):
        messagebox.showinfo(
            "About",
            "WooCommerce Product Sync\nVersion 1.0\n\nA tool for managing WooCommerce products and synchronizing with various vendors."
        )
    
    def test_connection(self, url=None, key=None, secret=None, settings_dialog=None):
        url = (url or self.url_input.get()).rstrip('/')
        key = key or self.key_input.get()
        secret = secret or self.secret_input.get()
        
        try:
            response = requests.get(
                f"{self.url_input.get()}/wp-json/wc/v3/products",
                auth=(self.key_input.get(), self.secret_input.get()),
                params={"per_page": 1}
            )
            
            if response.status_code == 200:
                status_text = "Successfully connected to WooCommerce!"
                status_color = "green"
                result = True
            else:
                status_text = f"Connection failed: {response.status_code}"
                status_color = "red"
                result = False
        except Exception as e:
            status_text = f"Error: {str(e)}"
            status_color = "red"
            result = False

        # Update status label in the appropriate window
        if settings_dialog:
            settings_dialog.status_label.config(text=status_text, foreground=status_color)
        else:
            self.status_label.config(text=status_text, foreground=status_color)
        
        return result

    def fetch_products(self):
        def fetch_thread():
            try:
                load_dotenv(override=True)
                url = os.getenv('WOO_API_URL', '').rstrip('/')
                key = os.getenv('WOO_API_KEY', '')
                secret = os.getenv('WOO_API_SECRET', '')
                print("Fetching products from API...")
                page = 1
                per_page = 50
                products_processed = 0
                product_queue = queue.Queue(maxsize=100)
                stop_event = threading.Event()
                
                # Get total number of products
                response = requests.get(
                    f"{url}/wp-json/wc/v3/products",
                    auth=(key, secret),
                    params={"per_page": 1}
                )
                total_items = int(response.headers.get('X-WP-Total', 0))
                progress_dialog = ProgressDialog(self, total_items)
                self.root.update_idletasks()
                
                def producer():
                    nonlocal page, products_processed
                    while not stop_event.is_set():
                        try:
                            response = requests.get(
                                f"{url}/wp-json/wc/v3/products",
                                auth=(key, secret),
                                params={"per_page": per_page, "page": page}
                            )
                            
                            if response.status_code == 200:
                                products_batch = response.json()
                                if not products_batch:
                                    product_queue.put(None)  # Signal consumer to stop
                                    break
                                
                                for product in products_batch:
                                    if stop_event.is_set():
                                        break
                                    product_queue.put(product)
                                page += 1
                            else:
                                print(f"Error fetching products: {response.status_code}")
                                break
                        except Exception as e:
                            print(f"Error in producer thread: {str(e)}")
                            break
                    product_queue.put(None)  # Ensure consumer stops
                
                def consumer():
                    nonlocal products_processed
                    while not stop_event.is_set():
                        try:
                            product = product_queue.get(timeout=5)  # 5 seconds timeout
                            if product is None:  # Stop signal
                                break
                            
                            try:
                                if not stop_event.is_set():
                                    self.db.add_or_update_product(product)
                                    products_processed += 1
                                progress_dialog.progress['value'] = products_processed
                                progress_dialog.label.config(text=f"Processing product {products_processed} of {total_items}")
                            except Exception as e:
                                print(f"Error processing product: {str(e)}")
                            finally:
                                product_queue.task_done()
                        except queue.Empty:
                            continue
                        except Exception as e:
                            print(f"Error in consumer thread: {str(e)}")
                            break
                
                # Start producer and consumer threads
                producer_thread = threading.Thread(target=producer)
                consumer_thread = threading.Thread(target=consumer)
                
                producer_thread.start()
                consumer_thread.start()
                
                # Monitor progress and check for cancellation
                while producer_thread.is_alive() or consumer_thread.is_alive():
                    if progress_dialog.is_cancelled:
                        stop_event.set()  # Signal threads to stop
                        break
                    
                    progress_dialog.update_progress(
                        products_processed,
                        f"Processed {products_processed} of {total_items} products"
                    )
                    time.sleep(0.1)
                
                # Wait for threads to finish
                producer_thread.join()
                consumer_thread.join()
                
                # Final update
                final_message = "Cancelled" if progress_dialog.is_cancelled else f"Completed! Processed {products_processed} products"
                progress_dialog.update_progress(products_processed, final_message)
                time.sleep(1)  # Show completion message briefly
                progress_dialog.dialog.destroy()
                
                # Refresh the product list
                self.update_product_list()
                
            except Exception as e:
                print(f"Error in fetch thread: {str(e)}")
                if 'progress_dialog' in locals():
                    progress_dialog.dialog.destroy()
        
        fetch_thread_instance = threading.Thread(target=fetch_thread)
        fetch_thread_instance.start()

class ProgressDialog:
    def __init__(self, app, total_items):
        self.dialog = tk.Toplevel(app.root)
        self.dialog.title("Fetching Products")
        self.dialog.geometry("400x180")
        self.dialog.transient(app.root)
        self.dialog.grab_set()
        self.tree = app.tree
        self.is_cancelled = False

        # Configure grid
        self.dialog.grid_columnconfigure(0, weight=1)
        self.dialog.grid_rowconfigure(2, weight=1)

        # Title label
        self.title_label = ttk.Label(self.dialog, text="Fetching Products", font=("Helvetica", 12, "bold"))
        self.title_label.grid(row=0, column=0, pady=(15, 5), padx=20)

        # Progress bar
        self.progress = ttk.Progressbar(self.dialog, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        self.progress['maximum'] = total_items

        # Status label
        self.label = ttk.Label(self.dialog, text="Starting...")
        self.label.grid(row=2, column=0, pady=10, padx=20)

        # Stop button
        self.stop_button = ttk.Button(self.dialog, text="Stop", command=self.stop_fetching)
        self.stop_button.grid(row=3, column=0, pady=(0, 15))

        # Handle window close button
        self.dialog.protocol("WM_DELETE_WINDOW", self.stop_fetching)

        # Update the dialog periodically
        self.dialog.after(100, self.periodic_update)

    def periodic_update(self):
        self.dialog.update_idletasks()
        if self.dialog.winfo_exists():
            self.dialog.after(100, self.periodic_update)

    def update_progress(self, current_item, message):
        self.progress['value'] = current_item
        self.label.config(text=message)
        self.dialog.update_idletasks()

    def stop_fetching(self):
        self.is_cancelled = True
        self.label.config(text="Stopping...")
        self.stop_button.config(state=tk.DISABLED)

    def close(self):
        if self.dialog.winfo_exists():
            self.dialog.destroy()

    def sort_treeview(self, col):
        # Get all items in the treeview
        items = [(self.tree.set(item, col), item) for item in self.tree.get_children('')]
        
        if not items:
            return
        
        # Determine if the column contains numeric values
        try:
            # Try to convert the first value to float
            float(items[0][0]) if items[0][0] != 'N/A' else 0
            is_numeric = True
        except ValueError:
            is_numeric = False
        
        # Sort items
        items.sort(key=lambda x: float(x[0]) if is_numeric and x[0] != 'N/A' else (float('-inf') if x[0] == 'N/A' else x[0].lower()),
                   reverse=getattr(self, 'sort_reverse', False))
        
        # Rearrange items in sorted positions
        for index, (val, item) in enumerate(items):
            self.tree.move(item, '', index)
            # Maintain link style for ID column
            if col != 'ID' and self.tree.set(item, 'ID') != 'N/A':
                self.tree.item(item, tags=('link',))
        
        # Reverse sort next time
        self.sort_reverse = not getattr(self, 'sort_reverse', False)
        
        # Update column headers to show sort direction
        for column in self.tree['columns']:
            self.tree.heading(column, text=column)
        self.tree.heading(col, text=f"{col} {'↓' if self.sort_reverse else '↑'}")

def main():
    root = tk.Tk()
    app = SyncApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()


    def on_double_click(self, event):
        # Get the item and column that was clicked
        region = self.tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
            
        column = self.tree.identify_column(event.x)
        item = self.tree.identify('item', event.x, event.y)
        
        if not item:
            return
            
        # Get column name
        col_id = int(column.replace('#', '')) - 1
        col_name = self.tree['columns'][col_id]
        
        # Only allow editing of price and stock columns
        if col_name not in ('Regular Price', 'Sale Price', 'Stock'):
            return
            
        # Get the current value and position
        current_value = self.tree.item(item)['values'][col_id]
        if current_value == 'N/A':
            current_value = ''
            
        # Get the bbox of the cell
        bbox = self.tree.bbox(item, column)
        if not bbox:
            return
            
        # Create an entry widget for editing
        entry = ttk.Entry(self.tree, width=len(str(current_value)) + 5)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        
        # Position the entry widget
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        
        def on_entry_return(event):
            self.save_edit(item, col_name, entry.get())
            entry.destroy()
            
        def on_entry_escape(event):
            entry.destroy()
            
        def on_focus_out(event):
            entry.destroy()
            
        entry.bind('<Return>', on_entry_return)
        entry.bind('<Escape>', on_entry_escape)
        entry.bind('<FocusOut>', on_focus_out)
        entry.focus_set()
        
        # Store reference to prevent garbage collection
        self.edit_widget = entry
        
    def save_edit(self, item, column, value):
        try:
            # Get the product ID from the tree item
            product_id = self.tree.item(item)['values'][0]
            if product_id == 'N/A':
                return
                
            # Validate and convert input
            if value.strip() == '':
                value = None
            elif column in ('Regular Price', 'Sale Price'):
                try:
                    value = float(value)
                    if self.show_tva_var.get():
                        value = value / 1.19  # Remove TVA for storage
                except ValueError:
                    messagebox.showerror('Invalid Input', 'Please enter a valid number for price')
                    return
            elif column == 'Stock':
                try:
                    value = int(value)
                except ValueError:
                    messagebox.showerror('Invalid Input', 'Please enter a valid number for stock')
                    return
                    
            # Update database
            column_map = {
                'Regular Price': 'regular_price',
                'Sale Price': 'sale_price',
                'Stock': 'stock_quantity'
            }
            
            db_column = column_map[column]
            self.db.update_product_field(product_id, db_column, value)
            
            # Update the tree
            values = list(self.tree.item(item)['values'])
            col_idx = self.tree['columns'].index(column)
            
            # Format the display value
            if column in ('Regular Price', 'Sale Price') and value is not None:
                if self.show_tva_var.get():
                    value = value * 1.19
                values[col_idx] = f"{value:.2f}"
            else:
                values[col_idx] = str(value) if value is not None else 'N/A'
                
            self.tree.item(item, values=values)
            
            # Show success message
            self.status_label.config(text=f"{column} updated successfully", foreground="green")
            
        except Exception as e:
            messagebox.showerror('Error', f'Failed to update {column.lower()}: {str(e)}')