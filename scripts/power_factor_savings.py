import json
import logging
import math
import sys
from pathlib import Path
from typing import Dict, List, Tuple
import csv
import xlsxwriter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class CalculatePowerFactorSavings:
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
      logging.error(f"{self.config_file} does not exist. Exiting.")
      sys.exit()
    logging.info(f"{self.config_file} exists. Loading...")
    with open(config_path) as f:
      config = json.load(f)
      logging.info(f"Loaded {self.config_file}.")
    return config

  def calculate_power_factor_savings(self) -> None:
    """Calculate power factor savings and write to an XLSX file."""
    year = self.config["year"]
    site = self.config["site"]
    results_dir = Path('results')
    results_dir.mkdir(exist_ok=True)
    file_name = results_dir / f"{year}_{site}_Power_Factor_Savings.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    logging.info(f"Created XLSX file: {file_name}")

    summary_worksheet = workbook.add_worksheet(f"{site}_{year}_Savings_Summary")
    logging.info(f"Created summary worksheet: {site}_{year}_Savings_Summary.")

    for month in self.months_list:
      self._process_month_data(workbook, month)

    self._calculate_year_power_factor_savings(workbook)
    workbook.close()
    logging.info(f"Completed calculating power factor savings and writing: {file_name}")

  def _process_month_data(self, workbook: xlsxwriter.Workbook, month: str) -> None:
    """Process data for a specific month and write to the workbook."""
    year = self.config["year"]
    site = self.config["site"]
    worksheet = workbook.add_worksheet(f"{site}_{month}_Savings")
    file_path = Path(f"data/power_factor_data/{year}/{month}-Power-Factor-Data.csv")
    if not file_path.exists():
      logging.warning(f"{file_path} does not exist. Skipping {month}.")
      return

    with open(file_path, mode='r', newline='') as infile:
      reader = list(csv.reader(infile))
      self._write_calculations_to_worksheet(worksheet, reader, month)

  def _write_calculations_to_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]], month: str) -> None:
    """Write calculations to the worksheet."""
    power_factor_col_idx = self._get_column_index(reader[0], "Power Factor")
    if power_factor_col_idx is None:
      logging.error(f"Power Factor column not found in {month} data. Skipping.")
      return

    total_cost_savings = self._calculate_savings(reader, power_factor_col_idx)
    self._write_savings_to_worksheet(worksheet, reader, total_cost_savings, month)
    self._adjust_column_widths(worksheet, reader)

  def _get_column_index(self, header: List[str], column_name: str) -> int:
    """Get the index of a column in the header."""
    for col_idx, col_name in enumerate(header):
      if col_name == column_name:
        return col_idx
    return None

  def _calculate_savings(self, reader: List[List[str]], power_factor_col_idx: int) -> float:
    """Calculate total cost savings."""
    total_cost_savings = 0.0
    for row in reader[1:]:
      try:
        power_factor = float(row[power_factor_col_idx])
        if power_factor < self.config["power_factor"]["target_power_factor"]:
          penalty = self.config["power_factor"]["penalty_rate"] * (self.config["power_factor"]["target_power_factor"] - power_factor)
          total_cost_savings += penalty
      except ValueError:
        continue
    return total_cost_savings

  def _write_savings_to_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]], total_cost_savings: float, month: str) -> None:
    """Write savings data to the worksheet."""
    worksheet.write(len(reader) + 1, 0, "Total Cost Savings")
    worksheet.write(len(reader) + 1, 1, total_cost_savings)
    logging.info(f"Calculated power factor savings for {month}.")

  def _adjust_column_widths(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]]) -> None:
    """Adjust the column widths based on the content."""
    for col_idx, col_name in enumerate(reader[0]):
      max_len = max(len(str(cell)) for cell in [col_name] + [row[col_idx] for row in reader[1:]])
      worksheet.set_column(col_idx, col_idx, max_len)
    logging.info(f"Adjusted column widths for worksheet {worksheet.name}.")

  def _calculate_year_power_factor_savings(self, workbook: xlsxwriter.Workbook) -> None:
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
    total_cost_savings_year = 0.0

    for row_idx, month in enumerate(self.months_list, start=1):
      month_data = self._read_month_data(month)
      total_cost_savings_year += month_data["total_cost_savings"]
      self._write_month_summary(summary_worksheet, row_idx, month, month_data)

    self._write_yearly_totals(summary_worksheet, total_cost_savings_year)

  def _read_month_data(self, month: str) -> Dict[str, float]:
    """Read month data from the worksheet."""
    year = self.config["year"]
    site = self.config["site"]
    file_path = Path(f"results/{year}_{site}_Power_Factor_Savings.xlsx")
    worksheet_name = f"{site}_{month}_Savings"

    try:
      workbook = xlsxwriter.Workbook(file_path)
      worksheet = workbook.get_worksheet_by_name(worksheet_name)
      total_cost_savings = worksheet.cell(len(reader) + 2, 2).value
      return {"total_cost_savings": total_cost_savings}
    except FileNotFoundError:
      logging.error(f"{file_path} not found.")
      return {"total_cost_savings": 0.0}
    except KeyError:
      logging.error(f"Worksheet {worksheet_name} not found in {file_path}.")
      return {"total_cost_savings": 0.0}
    except Exception as e:
      logging.error(f"Error reading month data: {e}")
      return {"total_cost_savings": 0.0}

  def _write_month_summary(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class, row_idx: int, month: str, month_data: Dict[str, float]) -> None:
    """Write the summary for a specific month."""
    summary_worksheet.write(row_idx, 0, month)
    summary_worksheet.write(row_idx, 1, month_data["total_cost_savings"])

  def _write_yearly_totals(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class, total_cost_savings_year: float) -> None:
    """Write the yearly totals to the summary worksheet."""
    row_idx = len(self.months_list) + 1
    summary_worksheet.write(row_idx, 0, "Yearly Totals")
    summary_worksheet.write(row_idx, 1, total_cost_savings_year)
    logging.info("Written yearly totals to summary worksheet.")

if __name__ == "__main__":
  config_file = "config/savings_config.json"
  calculator = CalculatePowerFactorSavings(config_file)
  calculator.calculate_power_factor_savings()