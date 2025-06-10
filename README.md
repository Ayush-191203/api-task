# Excel Processing API

A FastAPI-based REST API for processing Excel files and extracting table data.

## Features

- **List Tables**: Get all table names from the Excel file
- **Table Details**: Get row names for a specific table
- **Row Sum**: Calculate sum of numerical values in a specific row

## API Endpoints

### Base URL: `http://localhost:9090`

### Endpoints:

1. **GET /list_tables**
   - Returns all table names from the Excel file
   - Response: `{"tables": ["Initial Investment", "Revenue Projections", ...]}`

2. **GET /get_table_details?table_name={name}**
   - Returns row names for the specified table
   - Response: `{"table_name": "...", "row_names": [...]}`

3. **GET /row_sum?table_name={name}&row_name={name}**
   - Returns sum of numerical values in the specified row
   - Response: `{"table_name": "...", "row_name": "...", "sum": 123}`

## Setup and Installation

### Prerequisites
- Python 3.8+
- pip

### Installation Steps

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd APITASK
   ```

2. **Create virtual environment**:
   ```bash
   python -m venv venv
   ```

3. **Activate virtual environment**:
   - **Windows**:
     ```bash
     venv\Scripts\activate
     ```
   - **macOS/Linux**:
     ```bash
     source venv/bin/activate
     ```

4. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

5. **Run the application**:
   ```bash
   python main.py
   ```

## File Structure

```
APITASK/
├── venv/                 # Virtual environment
├── Data/
│   └── capbudg.xls      # Excel data file
├── Include/             # Virtual environment includes
├── Lib/                 # Virtual environment libraries
├── Scripts/             # Virtual environment scripts
├── main.py              # Main FastAPI application
├── pyvenv.cfg           # Virtual environment config
├── requirements.txt     # Python dependencies
├── .gitignore          # Git ignore rules
└── README.md           # This file
```

## Usage Examples

### Using curl:

```bash
# List all tables
curl http://localhost:9090/list_tables

# Get table details
curl "http://localhost:9090/get_table_details?table_name=Initial Investment"

# Get row sum
curl "http://localhost:9090/row_sum?table_name=Initial Investment&row_name=Tax Credit (if any )="
```

### Using browser:
Navigate to `http://localhost:9090/docs` for interactive API documentation.

## Development

### Debug Endpoints:
- `/debug` - View system information and loaded data
- `/reload_data` - Reload Excel data without restarting server


## Troubleshooting

1. **File not found error**: Ensure `capbudg.xls` is in the `Data/` directory
2. **Import errors**: Make sure all dependencies are installed via `pip install -r requirements.txt`
3. **Port already in use**: Change port in `main.py` or kill existing process

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License.
