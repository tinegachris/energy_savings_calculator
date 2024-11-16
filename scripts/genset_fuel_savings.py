import json
import csv
from pathlib import Path
import os
import sys
from typing import Dict, List, Tuple, Optional
import xlsxwriter
import openpyxl
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

    summary_worksheet = workbook.add_worksheet(f"{site}_{year}_Savings_Summary")
    logging.info(f"Created summary worksheet: {site}_{year}_Savings_Summary.")

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
    self._calculate_year_genset_savings(workbook, reader)
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
    total_month_genset_savings = genset_fuel_savings + outage_savings
    worksheet.write(len(reader) + 1, 0, "Total kWh Saved")
    worksheet.write(len(reader) + 1, 1, total_kwh_saved)
    worksheet.write(len(reader) + 3, 0, "Number of Outages")
    worksheet.write(len(reader) + 3, 1, num_outages)
    worksheet.write(len(reader) + 5, 0, "Genset Fuel Savings")
    worksheet.write(len(reader) + 5, 1, genset_fuel_savings)
    worksheet.write(len(reader) + 7, 0, "Outage Savings")
    worksheet.write(len(reader) + 7, 1, outage_savings)
    worksheet.write(len(reader) + 9, 0, "Total Month Genset Savings")
    worksheet.write(len(reader) + 9, 1, total_month_genset_savings)
    worksheet.write(len(reader) + 11, 0, "Solar Yield")
    worksheet.write(len(reader) + 11, 1, self.solar_yield_data.get(month, 0))
    worksheet.write(len(reader) + 13, 0, "Grid Yield")
    worksheet.write(len(reader) + 13, 1, self.grid_yield_data.get(month, 0))
    worksheet.write(len(reader) + 15, 0, "Genset Yield")
    worksheet.write(len(reader) + 15, 1, self.genset_yield_data.get(month, 0))
    logging.info(f"Calculated genset savings for {month}.")

    for col_idx, col_name in enumerate(reader[0]):
      max_len = max(len(str(cell)) for cell in [col_name] + [row[col_idx] for row in reader[1:]])
      worksheet.set_column(col_idx, col_idx, max_len + 2)
    logging.info(f"Calculated genset savings for {month} and written to the worksheet.")

  def _calculate_year_genset_savings(self, workbook: xlsxwriter.Workbook, reader: List[List[str]]) -> None:
    """Calculate yearly savings and update the summary worksheet."""
    site = self.config["site"]
    year = self.config["year"]
    summary_worksheet = workbook.get_worksheet_by_name(f"{self.config['site']}_{self.config['year']}_Savings_Summary")
    logging.info(f"Opened summary worksheet: {site}_{year}_Savings_Summary.")

    summary_worksheet.write(0, 0, "Month")
    summary_worksheet.write(0, 1, "Total kWh Saved")
    summary_worksheet.write(0, 2, "Number of Outages")
    summary_worksheet.write(0, 3, "Genset Fuel Savings")
    summary_worksheet.write(0, 4, "Outage Savings")
    summary_worksheet.write(0, 5, "Total Month Genset Savings")
    summary_worksheet.write(0, 6, "Solar Yield")
    summary_worksheet.write(0, 7, "Grid Yield")
    summary_worksheet.write(0, 8, "Genset Yield")

    logging.info(f"Started calculating yearly savings and writing to summary worksheet.")

    total_kwh_saved_year = 0
    total_num_outages_year = 0
    total_genset_fuel_savings_year = 0
    total_outage_savings_year = 0
    total_solar_yield_year = 0
    total_grid_yield_year = 0
    total_genset_yield_year = 0
    total_genset_savings_year = 0

    for row_idx, month in enumerate(self.months_list, start=1):
      try:
        worksheet_name = f"{site}_{month}_Savings"
        worksheet_path = f"results/{year}_{site}_Genset_Fuel_Savings.xlsx"
        wb = openpyxl.load_workbook(worksheet_path, data_only=True)
        worksheet = wb[worksheet_name]
        total_kwh_saved = worksheet.cell(row=len(reader) + 2, column=2).value
        num_outages = worksheet.cell(row=len(reader) + 4, column=2).value
        genset_fuel_savings = worksheet.cell(row=len(reader) + 6, column=2).value
        outage_savings = worksheet.cell(row=len(reader) + 8, column=2).value
        total_genset_month_savings = worksheet.cell(row=len(reader) + 10, column=2).value
        solar_yield = worksheet.cell(row=len(reader) + 12, column=2).value
        grid_yield = worksheet.cell(row=len(reader) + 14, column=2).value
        genset_yield = worksheet.cell(row=len(reader) + 16, column=2).value
        logging.info(f"Read yearly savings for {month} from the worksheet.")
      except FileNotFoundError:
        logging.error(f"File {worksheet_path} not found. Skipping {month}.")
        continue
      except KeyError:
        logging.error(f"Worksheet {worksheet_name} not found in {worksheet_path}. Skipping {month}.")
        continue
      except Exception as e:
        logging.error(f"An error occurred while processing {month}: {e}")
        continue

      summary_worksheet.write(row_idx, 0, month)
      summary_worksheet.write(row_idx, 1, total_kwh_saved)
      summary_worksheet.write(row_idx, 2, num_outages)
      summary_worksheet.write(row_idx, 3, genset_fuel_savings)
      summary_worksheet.write(row_idx, 4, outage_savings)
      summary_worksheet.write(row_idx, 5, total_genset_month_savings)
      summary_worksheet.write(row_idx, 6, solar_yield)
      summary_worksheet.write(row_idx, 7, grid_yield)
      summary_worksheet.write(row_idx, 8, genset_yield)
      logging.info(f"Written yearly savings for {month} to the summary worksheet.")

      try:
        total_kwh_saved_year += float(total_kwh_saved) if total_kwh_saved is not None else 0
        total_num_outages_year += int(num_outages) if num_outages is not None else 0
        total_genset_fuel_savings_year += float(genset_fuel_savings) if genset_fuel_savings is not None else 0
        total_outage_savings_year += float(outage_savings) if outage_savings is not None else 0
        total_genset_savings_year += float(total_genset_month_savings) if total_genset_month_savings is not None else 0
        total_solar_yield_year += float(solar_yield) if solar_yield is not None else 0
        total_grid_yield_year += float(grid_yield) if grid_yield is not None else 0
        total_genset_yield_year += float(genset_yield) if genset_yield is not None else 0
        logging.info(f"Updated yearly savings for {month}.")
      except ValueError as e:
        logging.error(f"Value error while updating yearly savings for {month}: {e}")
      except TypeError as e:
        logging.error(f"Type error while updating yearly savings for {month}: {e}")
      except Exception as e:
        logging.error(f"Unexpected error while updating yearly savings for {month}: {e}")

    summary_worksheet.write(len(self.months_list) + 1, 0, "Yearly Totals")
    summary_worksheet.write(len(self.months_list) + 1, 1, total_kwh_saved_year)
    summary_worksheet.write(len(self.months_list) + 1, 2, total_num_outages_year)
    summary_worksheet.write(len(self.months_list) + 1, 3, total_genset_fuel_savings_year)
    summary_worksheet.write(len(self.months_list) + 1, 4, total_outage_savings_year)
    summary_worksheet.write(len(self.months_list) + 1, 5, total_genset_savings_year)
    summary_worksheet.write(len(self.months_list) + 1, 6, total_solar_yield_year)
    summary_worksheet.write(len(self.months_list) + 1, 7, total_grid_yield_year)
    summary_worksheet.write(len(self.months_list) + 1, 8, total_genset_yield_year)

    logging.info(f"Completed calculating yearly savings and writing to summary worksheet.")

if __name__ == "__main__":
  config_file = "config/savings_config.json"
  calculator = CalculateGensetSavings(config_file)
  calculator.confirm_year_folder()
  calculator.clean_month_csv_files()
  calculator.remove_short_time_entries()
  calculator.calculate_genset_savings()
