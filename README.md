# Genset Savings Calculator

[![Github license](https://img.shields.io/github/license/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/blob/main/LICENSE)
[![GitHub contributors](https://img.shields.io/github/contributors/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/graphs/contributors)
[![GitHub issues](https://img.shields.io/github/issues/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/issues)
[![GitHub pull-requests](https://img.shields.io/github/issues-pr/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/pulls)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=flat-square)](http://makeapullrequest.com)
[![GitHub watchers](https://img.shields.io/github/watchers/tinegachris/genset_savings_calculator.svg?style=social&label=Watch)](https://github.com/tinegachris/genset_savings_calculator/watchers)
[![GitHub forks](https://img.shields.io/github/forks/tinegachris/genset_savings_calculator.svg?style=social&label=Fork)](https://github.com/tinegachris/genset_savings_calculator/network/members)
[![GitHub stars](https://img.shields.io/github/stars/tinegachris/genset_savings_calculator.svg?style=social&label=Star)](https://github.com/tinegachris/genset_savings_calculator/stargazers)

This project calculates genset fuel savings from CSV files and generates a summary in an Excel file.

![A visually appealing header image representing fuel savings and renewable energy integration. The image should feature a clean energy concept with solar panels, a genset (generator) emitting minimal exhaust, and a graph overlay showing energy savings. The background includes a blue sky and greenery, symbolizing sustainability. Incorporate a professional design suitable for technical documentation, with a focus on clarity and innovation.](Savings_Calculator_Header_Image.jpg)

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
