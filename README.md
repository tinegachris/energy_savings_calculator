# Genset Savings Calculator

This project calculates genset fuel savings from CSV files and generates a summary in an Excel file.

## Project Structure

- `scripts/genset_fuel_savings.py`: Main script for calculating genset savings.
- `config/savings_config.json`: Configuration file containing year, site, and cost details.
- `data/genset_savings_data/`: Directory containing monthly genset savings CSV files.
- `data/yield_data/`: Directory containing yield data CSV files.
- `results/`: Directory where the output Excel file will be saved.

## Requirements

- Python 3.x
- `xlsxwriter` library
- `openpyxl` library

## Installation

1. Clone the repository:

  ```sh
  git clone https://github.com/yourusername/genset_savings_calculator.git
  cd genset_savings_calculator
  ```

2. Install the required libraries:

  ```sh
  pip install -r requirements.txt
  ```

## Usage

1. Update the configuration file `config/savings_config.json` with the appropriate details.
2. Place the monthly genset savings CSV files in the `data/genset_savings_data/` directory.
3. Place the yield data CSV files in the `data/yield_data/` directory.
4. Run the main script:

  ```sh
  python scripts/genset_fuel_savings.py
  ```

## Output

The output will be an Excel file saved in the `results/` directory, containing a summary of the genset fuel savings.

## License

This project is licensed under the BSD 2-Clause License. See the [LICENSE](LICENSE) file for details.
