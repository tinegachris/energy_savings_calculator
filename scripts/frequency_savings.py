import json
import logging
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import csv
import xlsxwriter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class CalculateFrequencySavings:
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

  def confirm_year_folder(self) -> None:
    """Confirm the existence of the year folder."""
    year_folder = Path(f"data/frequency_savings_data/{self.config['year']}")
    if not year_folder.exists():
      logging.error(f"data/frequency_savings_data/{self.config['year']} folder does not exist. Exiting.")
      sys.exit()
    logging.info(f"data/frequency_savings_data/{self.config['year']} folder exists.")

  def calculate_frequency_savings(self) -> None:
    """Calculate frequency savings and write to an XLSX file."""
    year = self.config["year"]
    site = self.config["site"]
    results_dir = Path('results')
    results_dir.mkdir(exist_ok=True)
    file_name = results_dir / f"{year}_{site}_Frequency_Savings.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    logging.info(f"Created XLSX file: {file_name}")

    summary_worksheet = workbook.add_worksheet(f"{site}_{year}_Savings_Summary")
    logging.info(f"Created summary worksheet: {site}_{year}_Savings_Summary.")

    for month in self.months_list:
      self._process_month_data(workbook, month)

    self._calculate_year_frequency_savings(workbook)
    workbook.close()
    logging.info(f"Completed calculating frequency savings and writing: {file_name}")

  def _process_month_data(self, workbook: xlsxwriter.Workbook, month: str) -> None:
    """Process data for a specific month and write to the workbook."""
    year = self.config["year"]
    site = self.config["site"]
    worksheet = workbook.add_worksheet(f"{site}_{month}_Savings")
    file_path = Path(f"data/frequency_savings_data/{year}/{month}-Frequency-Savings.csv")
    if not file_path.exists():
      logging.warning(f"{file_path} does not exist. Skipping {month}.")
      return

    with open(file_path, mode='r', newline='') as infile:
      reader = list(csv.reader(infile))
      self._write_calculations_to_worksheet(worksheet, reader, month)

  def _write_calculations_to_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]], month: str) -> None:
    """Write calculations to the worksheet."""
    frequency_deviation_col_idx = self._get_column_index(reader[0], "Frequency Deviation (Hz)")
    if frequency_deviation_col_idx is None:
      logging.error(f"Frequency Deviation column not found in {month}. Skipping.")
      return

    total_deviation_cost, maintenance_cost, downtime_cost, penalty_cost = self._calculate_costs(reader, frequency_deviation_col_idx)
    total_month_savings = total_deviation_cost + maintenance_cost + downtime_cost + penalty_cost
    self._write_savings_to_worksheet(worksheet, reader, total_deviation_cost, maintenance_cost, downtime_cost, penalty_cost, total_month_savings, month)
    self._adjust_column_widths(worksheet, reader)

  def _get_column_index(self, header: List[str], column_name: str) -> Optional[int]:
    """Get the index of a column in the header."""
    for col_idx, col_name in enumerate(header):
      if col_name == column_name:
        return col_idx
    return None

  def _calculate_costs(self, reader: List[List[str]], frequency_deviation_col_idx: int) -> Tuple[float, float, float, float]:
    """Calculate costs associated with frequency deviations."""
    total_deviation_cost = 0
    maintenance_cost = 0
    downtime_cost = 0
    penalty_cost = 0

    for row in reader[1:]:
      try:
        deviation = float(row[frequency_deviation_col_idx])
        if abs(deviation) > self.config["frequency_deviation"]["tolerable_deviation"]:
          total_deviation_cost += abs(deviation) * self.config["frequency_deviation"]["cost_per_Hz_deviation"]
          maintenance_cost += abs(deviation) * self.config["frequency_deviation"]["maintenance_cost_increase_per_Hz"]
          downtime_cost += abs(deviation) * self.config["frequency_deviation"]["downtime_cost_per_Hz"]
          penalty_cost += abs(deviation) * self.config["frequency_deviation"]["penalty_rate"]
      except ValueError:
        continue

    return total_deviation_cost, maintenance_cost, downtime_cost, penalty_cost

  def _write_savings_to_worksheet(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]], total_deviation_cost: float, maintenance_cost: float, downtime_cost: float, penalty_cost: float, total_month_savings: float, month: str) -> None:
    """Write savings data to the worksheet."""
    worksheet.write(len(reader) + 1, 0, "Total Deviation Cost")
    worksheet.write(len(reader) + 1, 1, total_deviation_cost)
    worksheet.write(len(reader) + 3, 0, "Maintenance Cost")
    worksheet.write(len(reader) + 3, 1, maintenance_cost)
    worksheet.write(len(reader) + 5, 0, "Downtime Cost")
    worksheet.write(len(reader) + 5, 1, downtime_cost)
    worksheet.write(len(reader) + 7, 0, "Penalty Cost")
    worksheet.write(len(reader) + 7, 1, penalty_cost)
    worksheet.write(len(reader) + 9, 0, "Total Month Savings")
    worksheet.write(len(reader) + 9, 1, total_month_savings)
    logging.info(f"Calculated frequency savings for {month}.")

  def _adjust_column_widths(self, worksheet: xlsxwriter.Workbook.worksheet_class, reader: List[List[str]]) -> None:
    """Adjust the column widths based on the content."""
    for col_idx, col_name in enumerate(reader[0]):
      max_width = max(len(str(cell)) for cell in [col_name] + [row[col_idx] for row in reader[1:]])
      worksheet.set_column(col_idx, col_idx, max_width)
    logging.info(f"Adjusted column widths for worksheet {worksheet.name}.")

  def _calculate_year_frequency_savings(self, workbook: xlsxwriter.Workbook) -> None:
    """Calculate yearly savings and update the summary worksheet."""
    site = self.config["site"]
    year = self.config["year"]
    summary_worksheet = workbook.get_worksheet_by_name(f"{site}_{year}_Savings_Summary")
    logging.info(f"Opened summary worksheet: {site}_{year}_Savings_Summary.")

    self._write_summary_headers(summary_worksheet)
    self._aggregate_yearly_savings(summary_worksheet)

  def _write_summary_headers(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class) -> None:
    """Write headers to the summary worksheet."""
    headers = ["Month", "Total Deviation Cost", "Maintenance Cost", "Downtime Cost", "Penalty Cost", "Total Month Savings"]
    for col_idx, header in enumerate(headers):
      summary_worksheet.write(0, col_idx, header)
    logging.info("Written headers to summary worksheet.")

  def _aggregate_yearly_savings(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class) -> None:
    """Aggregate yearly savings and write to the summary worksheet."""
    totals = {
      "total_deviation_cost_year": 0,
      "total_maintenance_cost_year": 0,
      "total_downtime_cost_year": 0,
      "total_penalty_cost_year": 0,
      "total_savings_year": 0
    }

    for row_idx, month in enumerate(self.months_list, start=1):
      self._process_month_summary(summary_worksheet, row_idx, month, totals)

    self._write_yearly_totals(summary_worksheet, totals)

  def _process_month_summary(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class, row_idx: int, month: str, totals: Dict[str, float]) -> None:
    """Process and write the summary for a specific month."""
    year = self.config["year"]
    site = self.config["site"]
    worksheet_path = f"results/{year}_{site}_Frequency_Savings.xlsx"
    worksheet_name = f"{site}_{month}_Savings"

    try:
      workbook = xlsxwriter.Workbook(worksheet_path)
      worksheet = workbook.get_worksheet_by_name(worksheet_name)
      if worksheet is None:
        raise KeyError(f"Worksheet {worksheet_name} not found.")
      month_data = self._read_month_data(worksheet)
      self._write_month_summary(summary_worksheet, row_idx, month, month_data)
      self._update_yearly_totals(totals, month_data)
    except FileNotFoundError:
      logging.error(f"File {worksheet_path} not found.")
    except KeyError as e:
      logging.error(e)
    except Exception as e:
      logging.error(f"An error occurred while processing {month}: {e}")

  def _read_month_data(self, worksheet: xlsxwriter.Workbook.worksheet_class) -> Dict[str, float]:
    """Read month data from the worksheet."""
    return {
      "total_deviation_cost": worksheet.cell(row=len(reader) + 2, column=2).value,
      "maintenance_cost": worksheet.cell(row=len(reader) + 4, column=2).value,
      "downtime_cost": worksheet.cell(row=len(reader) + 6, column=2).value,
      "penalty_cost": worksheet.cell(row=len(reader) + 8, column=2).value,
      "total_savings": worksheet.cell(row=len(reader) + 10, column=2).value
    }

  def _write_month_summary(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class, row_idx: int, month: str, month_data: Dict[str, float]) -> None:
    """Write the summary for a specific month."""
    summary_worksheet.write(row_idx, 0, month)
    summary_worksheet.write(row_idx, 1, month_data["total_deviation_cost"])
    summary_worksheet.write(row_idx, 2, month_data["maintenance_cost"])
    summary_worksheet.write(row_idx, 3, month_data["downtime_cost"])
    summary_worksheet.write(row_idx, 4, month_data["penalty_cost"])
    summary_worksheet.write(row_idx, 5, month_data["total_savings"])

  def _update_yearly_totals(self, totals: Dict[str, float], month_data: Dict[str, float]) -> None:
    """Update the yearly totals with the data from a specific month."""
    try:
      totals["total_deviation_cost_year"] += month_data["total_deviation_cost"]
      totals["total_maintenance_cost_year"] += month_data["maintenance_cost"]
      totals["total_downtime_cost_year"] += month_data["downtime_cost"]
      totals["total_penalty_cost_year"] += month_data["penalty_cost"]
      totals["total_savings_year"] += month_data["total_savings"]
    except ValueError as e:
      logging.error(f"Value error: {e}")
    except TypeError as e:
      logging.error(f"Type error: {e}")
    except Exception as e:
      logging.error(f"An error occurred while updating yearly totals: {e}")

  def _write_yearly_totals(self, summary_worksheet: xlsxwriter.Workbook.worksheet_class, totals: Dict[str, float]) -> None:
    """Write the yearly totals to the summary worksheet."""
    row_idx = len(self.months_list) + 1
    summary_worksheet.write(row_idx, 0, "Yearly Totals")
    summary_worksheet.write(row_idx, 1, totals["total_deviation_cost_year"])
    summary_worksheet.write(row_idx, 2, totals["total_maintenance_cost_year"])
    summary_worksheet.write(row_idx, 3, totals["total_downtime_cost_year"])
    summary_worksheet.write(row_idx, 4, totals["total_penalty_cost_year"])
    summary_worksheet.write(row_idx, 5, totals["total_savings_year"])
    logging.info("Written yearly totals to summary worksheet.")

if __name__ == "__main__":
  config_file = "config/savings_config.json"
  calculator = CalculateFrequencySavings(config_file)
  calculator.confirm_year_folder()
  calculator.calculate_frequency_savings()