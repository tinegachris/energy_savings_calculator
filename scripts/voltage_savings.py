import csv
import xlsxwriter
from pathlib import Path
import logging
import json
import sys

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class VoltageSavingsCalculator:
    def __init__(self, config_file: str):
        self.config_file = config_file
        self.config = self.load_config()
        self.months_list = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]

    def load_config(self) -> dict:
        """Load configuration from a JSON file."""
        config_path = Path(self.config_file)
        if not config_path.exists():
            logging.error(f"{self.config_file} does not exist. Exiting.")
            sys.exit()
        logging.info(f"{self.config_file} exists. Loading...")
        with open(config_path) as f:
            config = json.load(f)
            logging.info(f"Loaded {self.config_file}.")
        return config

    def calculate_voltage_savings(self) -> None:
        """Calculate voltage savings for each month and save results in an Excel workbook."""
        year = self.config["year"]
        site = self.config["site"]
        file_name = f"{year}_{site}_Voltage_Savings.xlsx"
        workbook = xlsxwriter.Workbook(file_name)
        logging.info(f"Created XLSX file: {file_name}")

        for month in self.months_list:
            worksheet = workbook.add_worksheet(f"{month}_{site}_Voltage_Savings")
            file_path = Path(f"{year}/{month}-Voltage-Timeseries.csv")
            if not file_path.exists():
                logging.warning(f"{month}-Voltage-Timeseries.csv does not exist. Skipping.")
                continue

            with open(file_path, mode='r', newline='') as infile:
                reader = csv.DictReader(infile)
                fieldnames = reader.fieldnames
                data = list(reader)

            total_energy_deviation = 0
            for row in data:
                grid_voltage = float(row["Grid Voltage"])
                load_power = float(row["Load Power"])
                load_voltage = float(row["Load Voltage"])
                nominal_voltage = self.config["site_voltage"]

                energy_deviation = (nominal_voltage - grid_voltage) * load_power / nominal_voltage
                total_energy_deviation += energy_deviation

            grid_rate = self.config["cost_per_kWh"]
            battery_rate = self.config["cost_fuel_plus_MTCE"]

            cost_grid = total_energy_deviation * grid_rate
            battery_cost = total_energy_deviation * battery_rate
            savings = cost_grid - battery_cost

            worksheet.write(0, 0, "Total Energy Deviation (kWh)")
            worksheet.write(0, 1, total_energy_deviation)
            worksheet.write(1, 0, "Cost Without Battery Support (Grid Only)")
            worksheet.write(1, 1, cost_grid)
            worksheet.write(2, 0, "Cost With Battery Support")
            worksheet.write(2, 1, battery_cost)
            worksheet.write(3, 0, "Total Savings")
            worksheet.write(3, 1, savings)

            logging.info(f"Calculated voltage savings for {month} and written to the worksheet.")

        workbook.close()
        logging.info(f"Completed calculating voltage savings and writing: {file_name}")

if __name__ == "__main__":
    config_file = "config/savings_config.json"
    voltage_savings_calculator = VoltageSavingsCalculator(config_file)
    voltage_savings_calculator.calculate_voltage_savings()