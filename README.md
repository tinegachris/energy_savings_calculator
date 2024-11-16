# Energy Savings Calculator

[![Github license](https://img.shields.io/github/license/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/blob/main/LICENSE)
[![GitHub contributors](https://img.shields.io/github/contributors/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/graphs/contributors)
[![GitHub issues](https://img.shields.io/github/issues/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/issues)
[![GitHub pull-requests](https://img.shields.io/github/issues-pr/tinegachris/genset_savings_calculator.svg)](https://github.com/tinegachris/genset_savings_calculator/pulls)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=flat-square)](http://makeapullrequest.com)
[![GitHub watchers](https://img.shields.io/github/watchers/tinegachris/genset_savings_calculator.svg?style=social&label=Watch)](https://github.com/tinegachris/genset_savings_calculator/watchers)
[![GitHub forks](https://img.shields.io/github/forks/tinegachris/genset_savings_calculator.svg?style=social&label=Fork)](https://github.com/tinegachris/genset_savings_calculator/network/members)
[![GitHub stars](https://img.shields.io/github/stars/tinegachris/genset_savings_calculator.svg?style=social&label=Star)](https://github.com/tinegachris/genset_savings_calculator/stargazers)

![A visually appealing header image representing fuel savings and renewable energy integration. The image should feature a clean energy concept with solar panels, a genset (generator) emitting minimal exhaust, and a graph overlay showing energy savings. The background includes a blue sky and greenery, symbolizing sustainability. Incorporate a professional design suitable for technical documentation, with a focus on clarity and innovation.](Savings_Calculator_Header_Image.jpg)

## Project Overview

This project calculates energy savings from EMS data using CSV files and generates a summary in an Excel file. It processes monthly data to compute savings and aggregates yearly totals.

## Project Structure

- `scripts/`: Directory containing the scripts used for calculating savings from the data.
  - `genset_fuel_savings.py`: Script for calculating genset fuel savings.
  - `power_factor_savings.py`: Script for calculating power factor savings.
  - `frequency_savings.py`: Script for calculating frequency savings.
  - `voltage_stability_savings.py`: Script for calculating voltage stability savings.
  - `harmonics_savings.py`: Script for calculating harmonics savings.
- `config/savings_config.json`: Configuration file containing year, site, and cost details.
- `data/`: Directory containing the CSV data files.
  - `genset_savings_data/`: Directory containing genset savings data.
  - `power_factor_savings_data/`: Directory containing power factor savings data.
  - `frequency_data/`: Directory containing frequency data.
  - `voltage_data/`: Directory containing voltage data.
  - `harmonics_data/`: Directory containing harmonics data.
  - `yield_data/`: Directory containing yield data.

## Requirements

- Python 3.x
- A `requirements.txt` file is included in the repository, listing all the necessary Python libraries and dependencies needed to run the scripts.

## Installation

1. Clone the repository:

  ```sh
  git clone https://github.com/tinegachris/genset_savings_calculator.git
  cd genset_savings_calculator
  ```

2. Install the required libraries:

  ```sh
  pip install -r requirements.txt
  ```

  Ensure your Python environment is properly set up and activated before installing the dependencies.

  If you add new dependencies, update the `requirements.txt` file by running:

  ```sh
  pip freeze > requirements.txt
  ```

## Usage

1. Update the configuration file `config/savings_config.json` with the appropriate details.
2. Place the monthly savings CSV files in the `data/[savings_data]/` directory.
3. Run the main script:

  ```sh
  python scripts/genset_fuel_savings.py
  ```

## Output

The output will be an Excel file saved in the `results/` directory, containing a summary of the savings calculated.

## Contributing

If you have an idea for an improvement or have found a bug, please open an issue or submit a pull request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Make your changes.
4. Commit your changes (`git commit -am 'Add new feature'`).
5. Push to the branch (`git push origin feature-branch`).
6. Open a pull request.

## License

This project is licensed under the BSD 2-Clause License. See the [LICENSE](LICENSE) file for details.
