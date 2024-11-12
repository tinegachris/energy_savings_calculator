import json
import csv
from pathlib import Path
import sys
from typing import Dict, List

def load_config(config_file: str) -> Dict:
  """Load configuration from a JSON file."""

  if not Path(config_file).exists():
    print(f"{config_file} does not exist. Exiting.\n\n")
    sys.exit()
  else:
    print(f"{config_file} exists. Loading...\n\n")
    with open(config_file) as f:
      config = json.load(f)
      print(f"Loaded {config_file}.\n\n")
    return config

def confirm_year_folder(year: int) -> None:
  """Confirm the existence of the year folder."""

  if not Path(str(year)).exists():
    print(f"Year {year} folder does not exist. Exiting.\n\n")
    sys.exit()
  else:
    print(f"Year {year} folder exists.\n\n")

def refine_month_data(config: Dict) -> Dict[str, List[List[str]]]:
  """Check data for each month from CSV files and refine it."""

  months_list = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
  refined_data = {}
  for month in months_list:
    file_path = Path(f"{config['year']}/{month}-Genset-Savings.csv")
    if not file_path.exists():
      print(f"File {month}-Genset-Savings.csv does not exist.\n\n")
      refined_data[month] = []
    else:
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

def calculate_genset_savings() -> None:
  """Calculate genset savings."""
  pass

def main() -> None:
  """Main function to execute the script."""
  config_file = "config/savings_config.json"
  config = load_config(config_file)
  confirm_year_folder(config["year"])
  refine_month_data(config)
  calculate_genset_savings()

if __name__ == "__main__":
  main()