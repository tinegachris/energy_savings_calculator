import csv
import math
import json
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PowerFactorSavingsCalculator:
  def __init__(self, config_file: str):
    self.config_file = config_file
    self.config = self.load_config()
    self.data_dir = Path(self.config["data_directory"])
    self.target_pf = self.config["target_power_factor"]
    self.cost_reactive = self.config["cost_reactive"]
    self.system_loss_factor = self.config["system_loss_factor"]
    self.energy_cost = self.config["energy_cost"]

  def load_config(self) -> dict:
    """Load configuration from a JSON file."""
    config_path = Path(self.config_file)
    if not config_path.exists():
      logging.error(f"{self.config_file} does not exist. Exiting.")
      exit()
    logging.info(f"{self.config_file} exists. Loading...")
    with open(config_path) as f:
      config = json.load(f)
      logging.info(f"Loaded {self.config_file}.")
    return config

  def calculate_power_factor_savings(self, filename: str) -> None:
    """Calculate power factor savings from the CSV file."""
    file_path = self.data_dir / filename
    if not file_path.exists():
      logging.error(f"{file_path} does not exist. Exiting.")
      return

    total_cost_before = 0
    total_cost_after = 0
    total_energy_savings = 0

    with open(file_path, mode='r', newline='') as infile:
      reader = csv.reader(infile, delimiter='\t')
      for row in reader:
        timestamp, real_power = row
        real_power = float(real_power)
        apparent_power = real_power / self.target_pf
        reactive_power_before = math.sqrt(apparent_power**2 - real_power**2)
        reactive_power_after = reactive_power_before - self.config["reactive_power_reduction"]

        cost_before = self.cost_reactive * reactive_power_before
        cost_after = self.cost_reactive * reactive_power_after
        energy_savings = self.system_loss_factor * (reactive_power_before**2 - reactive_power_after**2)

        total_cost_before += cost_before
        total_cost_after += cost_after
        total_energy_savings += energy_savings

    cost_savings = total_cost_before - total_cost_after
    total_cost_savings = cost_savings + (total_energy_savings * self.energy_cost)

    logging.info(f"Total Cost Savings: {total_cost_savings}")
    logging.info(f"Total Energy Savings: {total_energy_savings}")

if __name__ == "__main__":
  config_file = "config/power_factor_savings_config.json"
  calculator = PowerFactorSavingsCalculator(config_file)
  calculator.calculate_power_factor_savings("October-Load-Active-Power-kW.csv")