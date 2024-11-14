# Calculating Savings from Voltage Correction Using Batteries

With timeseries data for grid and load voltages every 5 minutes, you can compute more precise savings by calculating the power and energy contributions of the batteries for each 5-minute interval. Here’s how you could approach it:

## Step 1: Calculate Voltage Deviation for Each Interval

1. **Compute Voltage Deviation**: For each 5-minute interval, calculate the deviation from nominal voltage ΔV as:

  ```plaintext
  ΔV = 400 V - V_grid(t)
  ```

  where V_grid(t) is the grid voltage at each 5-minute timestamp t.

2. **Determine Battery Contribution per Interval**: For each interval where ΔV ≠ 0, the battery compensates for this deviation to maintain the load voltage at 400V. Track these intervals as they represent moments when the battery is actively correcting voltage.

## Step 2: Calculate Power Supplied by Batteries for Each Interval

1. **Determine Load Power for Each Interval**: If you have the load power demand P_load(t) recorded, you can adjust this based on the voltage deviation.

2. **Calculate Battery Power Contribution**: For each 5-minute interval where a deviation exists, estimate the power contribution by the battery as:

  ```plaintext
  P_battery(t) = P_load(t) × (ΔV / 400)
  ```

3. **Convert to Energy Supplied in kWh**: Since each interval is 5 minutes (or 1/12 hours), the energy supplied by the battery during each interval t is:

  ```plaintext
  E_battery(t) = P_battery(t) × (1/12) hours
  ```

  Sum these energy values across all intervals to get the total energy supplied by the batteries over the entire time period:

  ```plaintext
  E_total = Σ E_battery(t)
  ```

## Step 3: Calculate Costs and Savings

1. **Calculate Cost Without Battery Support (Grid Only)**: Estimate the cost of maintaining stable voltage from the grid over the total deviation time:

  ```plaintext
  Cost_grid = E_total × Grid Rate per kWh
  ```

2. **Calculate Cost With Battery Support**: Calculate the energy cost of battery charging and discharging for voltage compensation:

  ```plaintext
  Battery Cost = E_total × Battery Charge/Discharge Rate per kWh
  ```

3. **Total Savings**: The total savings from using battery voltage correction instead of relying on the grid alone will be:

  ```plaintext
  Savings = Cost_grid - Battery Cost
  ```

### Example Calculation with Timeseries Data

Let’s say:
- Grid Rate per kWh: $0.10
- Battery Charge/Discharge Rate per kWh: $0.03
- Nominal Voltage: 400V

Suppose you have a timeseries of grid voltage, load power, and load voltage over a day. For a specific interval:

1. Grid Voltage: 380V
2. Load Power: 100 kW
3. Interval Duration: 5 minutes (or 1/12 hours)

**Calculations for This Interval**:

1. **Voltage Deviation**: ΔV = 400 - 380 = 20 V
2. **Battery Power Contribution**: 

  ```plaintext
  P_battery = 100 kW × (20 / 400) = 5 kW
  ```

3. **Energy Supplied by Battery for 5 Minutes**:

  ```plaintext
  E_battery = 5 kW × (1/12) = 0.4167 kWh
  ```

4. **Cost Without Battery Support**:

  ```plaintext
  Cost_grid = 0.4167 × 0.10 = 0.04167 USD
  ```

5. **Battery Cost**:

  ```plaintext
  Battery Cost = 0.4167 × 0.03 = 0.0125 USD
  ```

6. **Savings for This Interval**:

  ```plaintext
  Savings = 0.04167 - 0.0125 = 0.02917 USD
  ```


Sure, here is the text converted to markdown:
By summing the savings across all intervals in the day (or desired period), you can obtain the total savings due to voltage correction by the batteries.

With timeseries data for grid and load voltages every 5 minutes, you can compute more precise savings by calculating the power and energy contributions of the batteries for each 5-minute interval. Here’s how you could approach it:

## Step 1: Calculate Voltage Deviation for Each Interval

1. **Compute Voltage Deviation**: For each 5-minute interval, calculate the deviation from nominal voltage ΔV as:

  ```plaintext
  ΔV = 400 V - V_grid(t)
  ```

  where V_grid(t) is the grid voltage at each 5-minute timestamp t.

2. **Determine Battery Contribution per Interval**: For each interval where ΔV ≠ 0, the battery compensates for this deviation to maintain the load voltage at 400V. Track these intervals as they represent moments when the battery is actively correcting voltage.

## Step 2: Calculate Power Supplied by Batteries for Each Interval

1. **Determine Load Power for Each Interval**: If you have the load power demand P_load(t) recorded, you can adjust this based on the voltage deviation.

2. **Calculate Battery Power Contribution**: For each 5-minute interval where a deviation exists, estimate the power contribution by the battery as:

  ```plaintext
  P_battery(t) = P_load(t) × (ΔV / 400)
  ```

3. **Convert to Energy Supplied in kWh**: Since each interval is 5 minutes (or 1/12 hours), the energy supplied by the battery during each interval t is:

  ```plaintext
  E_battery(t) = P_battery(t) × (1/12) hours
  ```

  Sum these energy values across all intervals to get the total energy supplied by the batteries over the entire time period:

  ```plaintext
  E_total = Σ E_battery(t)
  ```

## Step 3: Calculate Costs and Savings

1. **Calculate Cost Without Battery Support (Grid Only)**: Estimate the cost of maintaining stable voltage from the grid over the total deviation time:

  ```plaintext
  Cost_grid = E_total × Grid Rate per kWh
  ```

2. **Calculate Cost With Battery Support**: Calculate the energy cost of battery charging and discharging for voltage compensation:

  ```plaintext
  Battery Cost = E_total × Battery Charge/Discharge Rate per kWh
  ```

3. **Total Savings**: The total savings from using battery voltage correction instead of relying on the grid alone will be:

  ```plaintext
  Savings = Cost_grid - Battery Cost
  ```

### Example Calculation with Timeseries Data

Let’s say:
- Grid Rate per kWh: $0.10
- Battery Charge/Discharge Rate per kWh: $0.03
- Nominal Voltage: 400V

Suppose you have a timeseries of grid voltage, load power, and load voltage over a day. For a specific interval:

1. Grid Voltage: 380V
2. Load Power: 100 kW
3. Interval Duration: 5 minutes (or 1/12 hours)

**Calculations for This Interval**:

1. **Voltage Deviation**: ΔV = 400 - 380 = 20 V
2. **Battery Power Contribution**: 

  ```plaintext
  P_battery = 100 kW × (20 / 400) = 5 kW
  ```

3. **Energy Supplied by Battery for 5 Minutes**:

  ```plaintext
  E_battery = 5 kW × (1/12) = 0.4167 kWh
  ```

4. **Cost Without Battery Support**:

  ```plaintext
  Cost_grid = 0.4167 × 0.10 = 0.04167 USD
  ```

5. **Battery Cost**:

  ```plaintext
  Battery Cost = 0.4167 × 0.03 = 0.0125 USD
  ```

6. **Savings for This Interval**:

  ```plaintext
  Savings = 0.04167 - 0.0125 = 0.02917 USD
  ```

By summing the savings across all intervals in the day (or desired period), you can obtain the total savings due to voltage correction by the batteries.
