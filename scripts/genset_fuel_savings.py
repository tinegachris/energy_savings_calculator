import json
import csv
from pathlib import Path
import os
import sys
from typing import Dict, List, Tuple, Optional
import xlsxwriter
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class CalculateGensetSavings:
  def __init__(self, config_file: str):
    self.config_file = config_file
    self.config = self.load_config()
    self.months_list = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ]
    self.solar_yield_data = self.load_yield_data("Solar-Energy-Yield-Month.csv", "Solar Energy Yield Month")
    self.grid_yield_data = self.load_yield_data("Grid-Energy-Yield-Month.csv", "Grid Energy Yield Month")
    self.genset_yield_data = self.load_yield_data("Genset-Energy-Yield-Month.csv", "Genset Energy Yield Month")

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

  def confirm_year_folder(self) -> None:
    """Confirm the existence of the year folder."""
    year_folder = Path(f"data/genset_savings_data/{self.config['year']}")
    if not year_folder.exists():
      logging.error(f"data/genset_savings_data/{self.config['year']} folder does not exist. Exiting.")
      sys.exit()
    logging.info(f"data/genset_savings_data/{self.config['year']} folder exists.")

  def clean_month_csv_files(self) -> None:
    """Remove the 'Conditions Met' column, remove previously calculated savings, and remove empty rows from the CSV file."""
    refined_data = {}
    for month in self.months_list:
      refined_month_data = []
      file_path = Path(f"data/genset_savings_data/{self.config['year']}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        logging.warning(f"data/genset_savings_data/{month}-Genset-Savings.csv does not exist.")
        continue
      with open(file_path, mode='r', newline='') as infile:
        reader = csv.DictReader(infile)
        fieldnames = [field for field in reader.fieldnames if field != "Conditions Met"]
        for row in reader:
          if any(row.values()) and "Total savings" not in row.values():
            refined_row = {field: row[field] for field in fieldnames}
            refined_month_data.append(refined_row)
      refined_data[month] = refined_month_data

    for month, data in refined_data.items():
      file_path = Path(f"data/genset_savings_data/{self.config['year']}/{month}-Genset-Savings.csv")
      with open(file_path, mode='w', newline='') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
      logging.info(f"Removed 'Conditions Met' column and empty rows from {file_path}.")

  def remove_short_time_entries(self) -> None:
    """Remove rows with 'Time Elapsed (minutes)' less than 1 from the CSV files."""
    for month in self.months_list:
      file_path = Path(f"data/genset_savings_data/{self.config['year']}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        logging.warning(f"{month}-Genset-Savings.csv does not exist.")
        continue
      filtered_data = []
      with open(file_path, mode='r', newline='') as infile:
        reader = csv.DictReader(infile)
        fieldnames = reader.fieldnames
        for row in reader:
          if float(row["Time Elapsed (minutes)"]) >= 1:
            filtered_data.append(row)
      with open(file_path, mode='w', newline='') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(filtered_data)
      logging.info(f"Removed rows with 'Time Elapsed (minutes)' < 1 from {file_path}.")

  def load_yield_data(self, filename: str, yield_column: str) -> Dict[str, str]:
    """Load yield data from CSV file."""
    year = self.config["year"]
    yield_file = Path(f"data/yield_data{year}/{filename}")
    if not yield_file.exists():
      logging.warning(f"{filename} does not exist. Skipping.")
      return {}
    with open(yield_file, mode='r', newline='') as infile:
      reader = csv.DictReader(infile)
      fieldnames = [field.strip() for field in reader.fieldnames]
      if 'Category' not in fieldnames or yield_column not in fieldnames:
        logging.error(f"Required columns not found in {yield_file}. Skipping.")
        return {}
      yield_data = {row["Category"]: row[yield_column] for row in reader}
    return yield_data

  def calculate_genset_savings(self) -> None:
    """Copy CSV data of each month into the XLSX file and calculate savings."""
    year = self.config["year"]
    site = self.config["site"]
    results_dir = Path('results')
    results_dir.mkdir(exist_ok=True)
    file_name = results_dir / f"{year}_{site}_Genset_Fuel_Savings.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    logging.info(f"Created XLSX file: {file_name}")

    for month in self.months_list:
      worksheet = workbook.add_worksheet(f"{site}_{month}_Savings")
      file_path = Path(f"data/genset_savings_data/{year}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        logging.warning(f"{file_path} does not exist. Skipping.")
        continue
      with open(file_path, mode='r', newline='') as infile:
        reader = list(csv.reader(infile))
        for row_idx, row in enumerate(reader):
          for col_idx, cell in enumerate(row):
            try:
              number = float(cell) if '.' in cell else int(cell)
              worksheet.write_number(row_idx, col_idx, number)
            except ValueError:
              worksheet.write(row_idx, col_idx, cell)
      logging.info(f"Copied data from {file_path} to {file_name} in sheet {site}_{month}_Savings.")
      self._write_calculations_to_worksheet(worksheet, reader, month)
    self._calculate_year_savings(workbook)
    workbook.close()
    logging.info(f"Completed calculating genset savings and writing: {file_name}")

  def _write_calculations_to_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]], month: str) -> None:
    """Write calculations to the worksheet."""
    energy_saving_col_idx = None
    for col_idx, col_name in enumerate(reader[0]):
      if col_name == "Energy Saving (kWh)":
        energy_saving_col_idx = col_idx
        break
    if energy_saving_col_idx is None:
      logging.warning(f"Energy Saving (kWh) column not found in {month}-Genset-Savings.csv. Skipping calculations.")
      return
    total_kwh_saved = sum(float(row[energy_saving_col_idx]) for row in reader[1:] if row[energy_saving_col_idx])
    num_outages = len([row for row in reader[1:] if row[energy_saving_col_idx]]) - 1
    cost_fuel_plus_MTCE = self.config["cost_fuel_plus_MTCE"]
    genset_fuel_savings = total_kwh_saved * cost_fuel_plus_MTCE
    cost_per_outage = self.config["cost_per_outage"]
    outage_savings = num_outages * cost_per_outage
    total_month_savings = genset_fuel_savings + outage_savings
    worksheet.write(len(reader) + 2, 0, "Total kWh Saved")
    worksheet.write(len(reader) + 2, 1, total_kwh_saved)
    worksheet.write(len(reader) + 4, 0, "Number of Outages")
    worksheet.write(len(reader) + 4, 1, num_outages)
    worksheet.write(len(reader) + 6, 0, "Genset Fuel Savings")
    worksheet.write(len(reader) + 6, 1, genset_fuel_savings)
    worksheet.write(len(reader) + 8, 0, "Outage Savings")
    worksheet.write(len(reader) + 8, 1, outage_savings)
    worksheet.write(len(reader) + 10, 0, "Solar Yield")
    worksheet.write(len(reader) + 10, 1, self.solar_yield_data.get(month, "N/A"))
    worksheet.write(len(reader) + 11, 0, "Grid Yield")
    worksheet.write(len(reader) + 11, 1, self.grid_yield_data.get(month, "N/A"))
    worksheet.write(len(reader) + 12, 0, "Genset Yield")
    worksheet.write(len(reader) + 12, 1, self.genset_yield_data.get(month, "N/A"))
    worksheet.write(len(reader) + 20, 0, "Total Month Savings")
    worksheet.write(len(reader) + 20, 1, total_month_savings)
    for col_idx, col_name in enumerate(reader[0]):
      max_len = max(len(str(cell)) for cell in [col_name] + [row[col_idx] for row in reader[1:]])
      worksheet.set_column(col_idx, col_idx, max_len + 2)
    logging.info(f"Calculated genset savings for {month} and written to the worksheet.")

  def _calculate_year_savings(self, workbook: xlsxwriter.Workbook) -> None:
    """Calculate yearly savings and update the summary worksheet."""
    year = self.config["year"]
    site = self.config["site"]
    summary_worksheet = workbook.add_worksheet(f"{site}_{year}_Savings_Summary")
    summary_worksheet.write(0, 0, "Month")
    summary_worksheet.write(0, 1, "Total kWh Saved")
    summary_worksheet.write(0, 2, "Number of Outages")
    summary_worksheet.write(0, 3, "Genset Fuel Savings")
    summary_worksheet.write(0, 4, "Outage Savings")
    summary_worksheet.write(0, 5, "Solar Yield")
    summary_worksheet.write(0, 6, "Grid Yield")
    summary_worksheet.write(0, 7, "Genset Yield")
    summary_worksheet.write(0, 8, "Total Month Savings")

    total_kwh_saved_year = 0
    total_num_outages_year = 0
    total_genset_fuel_savings_year = 0
    total_outage_savings_year = 0
    total_solar_yield_year = 0
    total_grid_yield_year = 0
    total_genset_yield_year = 0
    total_savings_year = 0

    for row_idx, month in enumerate(self.months_list, start=1):
      worksheet = workbook.get_worksheet_by_name(f"{site}_{month}_Savings")
      if worksheet is None:
        continue

      num_rows = worksheet.dim_rowmax + 1
      total_kwh_saved = worksheet.get_cell_value(num_rows + 2, 1)
      num_outages = worksheet.get_cell_value(num_rows + 4, 1)
      genset_fuel_savings = worksheet.get_cell_value(num_rows + 6, 1)
      outage_savings = worksheet.get_cell_value(num_rows + 8, 1)
      solar_yield = worksheet.get_cell_value(num_rows + 10, 1)
      grid_yield = worksheet.get_cell_value(num_rows + 11, 1)
      genset_yield = worksheet.get_cell_value(num_rows + 12, 1)
      total_month_savings = worksheet.get_cell_value(num_rows + 20, 1)

      summary_worksheet.write(row_idx, 0, month)
      summary_worksheet.write(row_idx, 1, total_kwh_saved)
      summary_worksheet.write(row_idx, 2, num_outages)
      summary_worksheet.write(row_idx, 3, genset_fuel_savings)
      summary_worksheet.write(row_idx, 4, outage_savings)
      summary_worksheet.write(row_idx, 5, solar_yield)
      summary_worksheet.write(row_idx, 6, grid_yield)
      summary_worksheet.write(row_idx, 7, genset_yield)
      summary_worksheet.write(row_idx, 8, total_month_savings)

      total_kwh_saved_year += total_kwh_saved
      total_num_outages_year += num_outages
      total_genset_fuel_savings_year += genset_fuel_savings
      total_outage_savings_year += outage_savings
      total_solar_yield_year += solar_yield
      total_grid_yield_year += grid_yield
      total_genset_yield_year += genset_yield
      total_savings_year += total_month_savings

    summary_worksheet.write(len(self.months_list) + 1, 0, "Yearly Totals")
    summary_worksheet.write(len(self.months_list) + 1, 1, total_kwh_saved_year)
    summary_worksheet.write(len(self.months_list) + 1, 2, total_num_outages_year)
    summary_worksheet.write(len(self.months_list) + 1, 3, total_genset_fuel_savings_year)
    summary_worksheet.write(len(self.months_list) + 1, 4, total_outage_savings_year)
    summary_worksheet.write(len(self.months_list) + 1, 5, total_solar_yield_year)
    summary_worksheet.write(len(self.months_list) + 1, 6, total_grid_yield_year)
    summary_worksheet.write(len(self.months_list) + 1, 7, total_genset_yield_year)
    summary_worksheet.write(len(self.months_list) + 1, 8, total_savings_year)

    logging.info(f"Completed calculating yearly savings and writing to summary worksheet.")

if __name__ == "__main__":
  config_file = "config/savings_config.json"
  genset_savings = CalculateGensetSavings(config_file)
  genset_savings.confirm_year_folder()
  genset_savings.clean_month_csv_files()
  genset_savings.remove_short_time_entries()
  genset_savings.calculate_genset_savings()
