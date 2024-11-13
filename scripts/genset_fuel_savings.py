import json
import csv
from pathlib import Path
import sys
from typing import Dict, List
import xlsxwriter

class CalculateGensetSavings:
  def __init__(self, config_file: str):
    self.config_file = config_file
    self.config = self.load_config()
    self.months_list = [
      "January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ]

  def load_config(self) -> Dict:
    """Load configuration from a JSON file."""
    config_path = Path(self.config_file)
    if not config_path.exists():
      print(f"{self.config_file} does not exist. Exiting.\n\n")
      sys.exit()
    print(f"{self.config_file} exists. Loading...\n\n")
    with open(config_path) as f:
      config = json.load(f)
      print(f"Loaded {self.config_file}.\n\n")
    return config

  def confirm_year_folder(self) -> None:
    """Confirm the existence of the year folder."""
    year_folder = Path(str(self.config["year"]))
    if not year_folder.exists():
      print(f"Year {self.config['year']} folder does not exist. Exiting.\n\n")
      sys.exit()
    print(f"Year {self.config['year']} folder exists.\n\n")

  def clean_month_csv_files(self) -> None:
    """Remove the 'Conditions Met' column, remove previously calculated savings, and remove empty rows from the CSV file."""
    refined_data = {}
    for month in self.months_list:
      refined_month_data = []
      file_path = Path(f"{self.config['year']}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        print(f"{month}-Genset-Savings.csv does not exist.\n\n")
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
      file_path = Path(f"{self.config['year']}/{month}-Genset-Savings.csv")
      with open(file_path, mode='w', newline='') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
      print(f"Removed 'Conditions Met' column and empty rows from {file_path}.\n\n")

  def remove_short_time_entries(self) -> None:
    """Remove rows with 'Time Elapsed (minutes)' less than 1 from the CSV files."""
    for month in self.months_list:
      file_path = Path(f"{self.config['year']}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        print(f"{month}-Genset-Savings.csv does not exist.\n\n")
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
      print(f"Removed rows with 'Time Elapsed (minutes)' < 1 from {file_path}.\n\n")

  def calculate_genset_savings(self) -> None:
    """Copy CSV data of each month into the XLSX file and calculate savings."""
    year = self.config["year"]
    site = self.config["site"]
    file_name = f"{year}_Genset_Fuel_Savings_{site}.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    print(f"Created XLSX file: {file_name}\n\n")
    summary_sheet = workbook.add_worksheet("Summary")

    for month in self.months_list:
      worksheet = workbook.add_worksheet(f"{site}_{month}_Savings")
      file_path = Path(f"{year}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        print(f"{month}-Genset-Savings.csv does not exist. Skipping.\n\n")
        continue
      with open(file_path, mode='r', newline='') as infile:
        reader = list(csv.reader(infile))
        for row_idx, row in enumerate(reader):
          for col_idx, cell in enumerate(row):
            worksheet.write(row_idx, col_idx, cell)
      print(f"Copied data from {file_path} to {file_name} in sheet {site}_{month}_Savings.\n\n")
      # Calculate total kWh Saved
      energy_saving_col_idx = None
      for col_idx, col_name in enumerate(reader[0]):
        if col_name == "Energy Saving (kWh)":
          energy_saving_col_idx = col_idx
          break
      if energy_saving_col_idx is None:
        print(f"Energy Saving (kWh) column not found in {file_path}. Skipping calculations.\n\n")
        continue
      total_kwh_saved = sum(float(row[energy_saving_col_idx]) for row in reader[1:] if row[energy_saving_col_idx])
      # Calculate number of outages
      num_outages = len([row for row in reader[1:] if row[energy_saving_col_idx]]) - 1
      # Calculate genset fuel savings
      fuel_plus_mtce_cost = self.config["fuel_plus_MTCE_cost"]
      genset_fuel_savings = total_kwh_saved * fuel_plus_mtce_cost
      # Calculate outage savings
      cost_per_outage = self.config["cost_per_outage"]
      outage_savings = num_outages * cost_per_outage
      # Write calculations to the worksheet
      worksheet.write(row_idx + 2, 0, "Total kWh Saved")
      worksheet.write(row_idx + 2, 1, total_kwh_saved)
      worksheet.write(row_idx + 3, 0, "Number of Outages")
      worksheet.write(row_idx + 3, 1, num_outages)
      worksheet.write(row_idx + 4, 0, "Genset Fuel Savings")
      worksheet.write(row_idx + 4, 1, genset_fuel_savings)
      worksheet.write(row_idx + 5, 0, "Outage Savings")
      worksheet.write(row_idx + 5, 1, outage_savings)

      for col_idx, col_name in enumerate(reader[0]):
        max_len = max(len(str(cell)) for cell in [col_name] + [row[col_idx] for row in reader[1:]])
        worksheet.set_column(col_idx, col_idx, max_len + 2)
    workbook.close()
    print(f"Completed calculating genset savings and writing: {file_name}\n\n")

if __name__ == "__main__":
  config_file = "config/savings_config.json"
  genset_savings = CalculateGensetSavings(config_file)
  genset_savings.confirm_year_folder()
  genset_savings.clean_month_csv_files()
  genset_savings.calculate_genset_savings()