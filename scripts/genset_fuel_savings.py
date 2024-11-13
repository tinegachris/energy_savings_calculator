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
    self.refined_data = self.refine_month_data()

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

  def refine_month_data(self) -> Dict[str, List[List[str]]]:
    """Check data for each month from CSV files and refine it."""
    refined_data = {}
    for month in self.months_list:
      file_path = Path(f"{self.config['year']}/{month}-Genset-Savings.csv")
      if not file_path.exists():
        print(f"File {month}-Genset-Savings.csv does not exist. Exiting.\n\n")
        sys.exit()
      month_data = []
      with open(file_path, newline='') as f:
        reader = csv.reader(f)
        headers = next(reader)
        if "Conditions Met" in headers:
          conditions_met_index = headers.index("Conditions Met")
          headers.pop(conditions_met_index)
        month_data.append(headers)
        for row in reader:
          if "Conditions Met" in headers:
            row.pop(conditions_met_index)
          month_data.append(row)
      refined_data[month] = month_data
      print(f"Processed {file_path}.\n\n")
    return refined_data

  def calculate_genset_savings(self) -> None:
    """Calculate genset savings."""
    pass

if __name__ == "__main__":
  config_file = "config\savings_config.json"
  genset_savings = CalculateGensetSavings(config_file)
  genset_savings.confirm_year_folder()
  genset_savings.calculate_genset_savings()