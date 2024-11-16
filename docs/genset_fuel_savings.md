# Genset Fuel Savings Calculation Guide

This guide explains how to calculate fuel savings from genset operations using monthly CSV data and configuration settings.

## Step 1: Load Configuration

1. **Load Configuration File:**

- The configuration file contains necessary parameters such as year, site, cost of fuel plus maintenance, and cost per outage.
- Example configuration file path: `config/savings_config.json`

## Step 2: Confirm Year Folder

1. **Check Year Folder Existence:**

- Ensure the folder for the specified year exists in the data directory.
- Example path: `data/genset_savings_data/{year}`

## Step 3: Clean Monthly CSV Files

1. **Remove Unnecessary Data:**

- Remove the 'Conditions Met' column.
- Remove previously calculated savings.
- Remove empty rows from the CSV files.

## Step 4: Remove Short Time Entries

1. **Filter Entries:**

- Remove rows where 'Time Elapsed (minutes)' is less than 1 minute.

## Step 5: Load Yield Data

1. **Load Yield Data from CSV Files:**

- Load solar, grid, and genset yield data from respective CSV files.
- Example paths:
  - `data/yield_data/{year}/Solar-Energy-Yield-Month.csv`
  - `data/yield_data/{year}/Grid-Energy-Yield-Month.csv`
  - `data/yield_data/{year}/Genset-Energy-Yield-Month.csv`

## Step 6: Calculate Genset Savings

1. **Create XLSX File:**

- Create an XLSX file to store the results.
- Example file name: `results/{year}_{site}_Genset_Fuel_Savings.xlsx`

2. **Process Monthly Data:**

- For each month, copy CSV data into the XLSX file and calculate savings.
- Calculate total kWh saved, number of outages, genset fuel savings, outage savings, and total month genset savings.

3. **Write Calculations to Worksheet:**

- Write the calculated savings data to the worksheet.
- Include solar yield, grid yield, and genset yield for each month.

## Step 7: Calculate Yearly Savings

1. **Aggregate Monthly Savings:**

- Aggregate the savings data for the entire year.
- Write the yearly totals to the summary worksheet.

## Step 8: Summary Worksheet

1. **Write Headers:**

- Write headers to the summary worksheet.
- Example headers: "Month", "Total kWh Saved", "Number of Outages", "Genset Fuel Savings", "Outage Savings", "Total Month Genset Savings", "Solar Yield", "Grid Yield", "Genset Yield"

2. **Write Monthly Summaries:**

- Write the summary for each month to the summary worksheet.

3. **Write Yearly Totals:**

- Write the yearly totals to the summary worksheet.

## Example Configuration File

```json
{
  "year": 2023,
  "site": "ExampleSite",
  "cost_of_fuel_plus_maintenance": 1.25,
  "cost_per_outage": 50.00
}
```