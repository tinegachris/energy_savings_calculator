import json
import logging
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import csv
import xlsxwriter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class CalculateVoltageStabilitySavings:
    def __init__(self, config_file: str):
        self.config_file = config_file
        self.config = self.load_config()
        self.months_list = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]
        self.grid_voltage_data = self.load_voltage_data("Grid-Voltage-Month.csv", "Grid_Voltage_Month")
        self.load_voltage_data = self.load_voltage_data("Load-Voltage-Month.csv", "Load_Voltage_Month")

    def load_config(self) -> Dict:
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

    def load_voltage_data(self, filename: str, voltage_column: str) -> Dict[str, float]:
        """Load voltage data from CSV file."""
        year = self.config["year"]
        voltage_file = Path(f"data/voltage_data/{year}/{filename}")
        if not voltage_file.exists():
            logging.warning(f"{filename} does not exist. Skipping.")
            return {}
        with open(voltage_file, mode='r', newline='') as infile:
            reader = csv.reader(infile)
            headers = next(reader)
            try:
                voltage_idx = headers.index(voltage_column)
            except ValueError:
                logging.error(f"{voltage_column} not found in {filename}.")
                return {}
            voltage_data = {}
            for row in reader:
                month = row[0]
                voltage = float(row[voltage_idx])
                voltage_data[month] = voltage
        return voltage_data

    def calculate_voltage_stability_savings(self) -> None:
        """Calculate voltage stability savings and write to an XLSX file."""
        year = self.config["year"]
        site = self.config["site"]
        results_dir = Path('results')
        results_dir.mkdir(exist_ok=True)
        file_name = results_dir / f"{year}_{site}_Voltage_Stability_Savings.xlsx"
        workbook = xlsxwriter.Workbook(file_name)
        logging.info(f"Created XLSX file: {file_name}")

        summary_worksheet = workbook.add_worksheet(f"{site}_{year}_Savings_Summary")
        logging.info(f"Created summary worksheet: {site}_{year}_Savings_Summary.")

        for month in self.months_list:
            self._process_month_data(workbook, month)

        self._calculate_year_voltage_savings(workbook)
        workbook.close()
        logging.info(f"Completed calculating voltage stability savings and writing: {file_name}")

    def _process_month_data(self, workbook: xlsxwriter.Workbook, month: str) -> None:
        """Process data for a specific month and write to the workbook."""
        year = self.config["year"]
        site = self.config["site"]
        worksheet = workbook.add_worksheet(f"{site}_{month}_Savings")
        grid_file_path = Path(f"data/voltage_data/{year}/{month}-Grid-Voltage.csv")
        load_file_path = Path(f"data/voltage_data/{year}/{month}-Load-Voltage.csv")
        if not grid_file_path.exists() or not load_file_path.exists():
            logging.warning(f"Voltage data files for {month} do not exist. Skipping.")
            return

        with open(grid_file_path, mode='r', newline='') as grid_infile, open(load_file_path, mode='r', newline='') as load_infile:
            grid_reader = csv.reader(grid_infile)
            load_reader = csv.reader(load_infile)
            grid_headers = next(grid_reader)
            load_headers = next(load_reader)
            grid_voltage_idx = grid_headers.index("Grid_Voltage")
            load_voltage_idx = load_headers.index("Load_Voltage")

            self._write_calculations_to_worksheet(worksheet, grid_reader, load_reader, grid_voltage_idx, load_voltage_idx, month)

    def _write_calculations_to_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class, grid_reader: List[List[str]], load_reader: List[List[str]], grid_voltage_idx: int, load_voltage_idx: int, month: str) -> None:
        """Write calculations to the worksheet."""
        nominal_voltage = self.config["voltage_stability"]["nominal_voltage"]
        voltage_tolerance = self.config["voltage_stability"]["voltage_tolerance"]
        cost_per_kWh_mismatch = self.config["voltage_stability"]["cost_per_kWh_mismatch"]
        efficiency_derate_per_percent_deviation = self.config["voltage_stability"]["efficiency_derate_per_percent_deviation"]

        total_cost_savings = 0
        for grid_row, load_row in zip(grid_reader, load_reader):
            grid_voltage = float(grid_row[grid_voltage_idx])
            load_voltage = float(load_row[load_voltage_idx])
            delta_v_grid = abs(grid_voltage - nominal_voltage)
            delta_v_load = abs(load_voltage - nominal_voltage)
            if delta_v_grid > voltage_tolerance * nominal_voltage:
                energy_mismatch = self.config["site_capacity"] * (1 / 12)  # Assuming 5-minute intervals
                cost_savings = energy_mismatch * cost_per_kWh_mismatch * (delta_v_grid - delta_v_load) / nominal_voltage
                total_cost_savings += cost_savings

        worksheet.write(0, 0, "Month")
        worksheet.write(0, 1, "Total Cost Savings")
        worksheet.write(1, 0, month)
        worksheet.write(1, 1, total_cost_savings)
        logging.info(f"Calculated voltage stability savings for {month}.")

    def _calculate_year_voltage_savings(self, workbook: xlsxwriter.Workbook) -> None:
        """Calculate yearly savings and update the summary worksheet."""
        site = self.config["site"]
        year = self.config["year"]
        summary_worksheet = workbook.get_worksheet_by_name(f"{site}_{year}_Savings_Summary")
        logging.info(f"Opened summary worksheet: {site}_{year}_Savings_Summary.")

        self._write_summary_headers(summary_worksheet)
        self._aggregate_yearly_savings(summary_worksheet)

    def _write_summary_headers(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class) -> None:
        """Write headers to the summary worksheet."""
        headers = ["Month", "Total Cost Savings"]
        for col_idx, header in enumerate(headers):
            summary_worksheet.write(0, col_idx, header)
        logging.info("Written headers to summary worksheet.")

    def _aggregate_yearly_savings(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class) -> None:
        """Aggregate yearly savings and write to the summary worksheet."""
        total_yearly_savings = 0
        for row_idx, month in enumerate(self.months_list, start=1):
            worksheet_name = f"{self.config['site']}_{month}_Savings"
            worksheet = summary_worksheet.book.get_worksheet_by_name(worksheet_name)
            if worksheet:
                total_cost_savings = worksheet.cell(1, 1).value
                summary_worksheet.write(row_idx, 0, month)
                summary_worksheet.write(row_idx, 1, total_cost_savings)
                total_yearly_savings += total_cost_savings

        summary_worksheet.write(len(self.months_list) + 1, 0, "Yearly Total")
        summary_worksheet.write(len(self.months_list) + 1, 1, total_yearly_savings)
        logging.info("Written yearly totals to summary worksheet.")

if __name__ == "__main__":
    config_file = "config/savings_config.json"
    calculator = CalculateVoltageStabilitySavings(config_file)
    calculator.calculate_voltage_stability_savings()