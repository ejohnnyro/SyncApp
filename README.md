# WooCommerce Product Sync

A Python-based desktop application for synchronizing WooCommerce products using PyQt6 for the user interface and the WooCommerce REST API for data management.

## Features

- Modern and responsive PyQt6-based user interface
- Secure WooCommerce API integration
- Product synchronization capabilities
- Easy-to-use configuration interface

## Requirements

- Python 3.8 or higher
- PyQt6 for the user interface
- Requests library for API communication
- python-dotenv for environment variable management

## Installation

1. Clone the repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Running the Application

Execute the following command in the project directory:

```
python sync_app.py
```

## Configuration

1. Launch the application
2. Enter your WooCommerce store URL
3. Input your WooCommerce API Key and Secret
4. Click "Connect to Store" to test the connection

## Security

API credentials are handled securely and are not stored in plain text. Always keep your API credentials confidential.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.