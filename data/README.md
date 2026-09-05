# ProcureFlow Data

## Official Dataset

ProcureFlow uses the **Aerospace Supply Chain Performance & Forecasting** dataset.

The official source files are:

- `parts_master.csv`
- `supply_chain_history.csv`
- `purchase_orders.csv`
- `quality_incidents.csv`

## Local Folder

Place the original source files in:

`data/raw/`

The raw CSV files are intentionally excluded from Git through `.gitignore`.

## Data Policy

The original source files must remain unchanged.

ProcureFlow will perform ingestion, typing, cleaning, validation and transformation through Power Query.

Manual modification of the raw CSV files is not part of the normal operating process.

## Sample Data

The folder `data/sample/` is reserved for optional lightweight samples if they are later justified for documentation or portfolio purposes.
