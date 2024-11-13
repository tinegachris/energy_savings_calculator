import json
import csv
from pathlib import Path
import sys
from typing import Dict, List

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

  def remove_conditions_met_column(self) -> None:
    """Remove the 'Conditions Met' column from the CSV file."""
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
          refined_row = {field: row[field] for field in fieldnames}
          refined_month_data.append(refined_row)
      refined_data[month] = refined_month_data

    for month, data in refined_data.items():
      file_path = Path(f"{self.config['year']}/{month}-Genset-Savings.csv")
      with open(file_path, mode='w', newline='') as outfile:
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
      print(f"Removed 'Conditions Met' column from {file_path}.\n\n")

  def calculate_genset_savings(self) -> None:
    """Calculate genset savings."""
    pass

if __name__ == "__main__":
  config_file = "config/savings_config.json"
  genset_savings = CalculateGensetSavings(config_file)
  genset_savings.confirm_year_folder()
  genset_savings.remove_conditions_met_column()