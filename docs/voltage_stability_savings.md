# Voltage Stability Savings Calculation Guide

To calculate cost savings from voltage stability improvements using your grid supply 5-minute time series and load voltage 5-minute time series, you can take the following approach:

## Step 1: Quantify Voltage Deviations

1. **Determine Voltage Deviation Deltas (ΔV):**

- For each timestamp in the data:
  - ΔV = |V_load - V_nominal|
  - V_load: Measured load voltage.
  - V_nominal: Nominal voltage (configurable e.g., 415V).

2. **Separate Positive and Negative Deviations:**

- Positive delta (+ΔV): Grid-supplied voltage was higher than nominal.
- Negative delta (-ΔV): Grid-supplied voltage was lower than nominal.

3. **Group Deviations by Severity:**

- Use voltage tolerance standards (e.g., ±5% for most utilities) to group deviations into acceptable and unacceptable categories.
- Define thresholds:
  - Unacceptable if |ΔV| > 0.05 × V_nominal
- Discard acceptable deviations

### A. Equipment Stress and Efficiency Losses

- Equipment operating outside acceptable voltage tolerances experiences:
  - Increased heat losses.
  - Reduced lifespan (transformers, motors, and sensitive electronics).
  - Malfunctioning control systems.

### B. Cost Estimation Approach

1. **Determine Energy Mismatch:**

- Energy affected by voltage instability:
  - E_ΔV = P_load × t
  - P_load: Load power during the time of deviation. (Load 5-minute time series)
  - t: Duration of the deviation (5 minutes = 300s = 1/12h).

2. **Associate Energy Mismatch with Cost:**

- Use cost implications for over/under-voltage from utility data or equipment loss data (e.g., derating or efficiency loss).

Example:

- Equipment derates efficiency by x% for every 1% deviation.

1. **Measure Corrected Voltage (V_BESS):**

- Compare grid voltage deviations (V_grid - V_nominal) with load voltage deviations after BESS correction (V_load - V_nominal).

2. **Quantify the Mitigation:**

- Reduction in voltage delta due to BESS correction:
  - ΔV_reduction = |ΔV_grid| - |ΔV_BESS|

3. **Associate Voltage Reduction with Cost Savings:**

- If BESS reduces deviations, calculate the avoided energy mismatch cost as:
  - Cost Savings = E_ΔV_reduction × C_impact
  - C_impact: Cost of inefficiency per unit energy mismatch (e.g., USD/kWh).
  - Reduction in voltage delta due to BESS correction:
    - ΔV_reduction = |ΔV_grid| - |ΔV_BESS|

3. **Associate Voltage Reduction with Cost Savings:**
1. **Aggregate cost savings over the entire month:**

- Total Cost Savings = Σ (Cost Savings per Time Interval)

2. **Compare two scenarios:**

- Without BESS: Based on grid voltage fluctuations (ΔV_grid).

## Additional Considerations

- **Data Source Validation:** Ensure accurate power and voltage measurements.
- **Voltage-Cost Mapping:** Obtain cost multipliers from utility penalties, maintenance logs, or equipment manufacturers.
- **Energy Tariffs:** Use real-time tariffs to enhance accuracy for savings calculations.

1. **Aggregate cost savings over the entire month:**

- Total Cost Savings = Σ (Cost Savings per Time Interval)

2. **Compare two scenarios:**

- Without BESS: Based on grid voltage fluctuations (ΔV_grid).

- With BESS: Based on corrected voltage (ΔV_BESS).

## Additional Considerations

- **Data Source Validation:** Ensure accurate power and voltage measurements.
- **Voltage-Cost Mapping:** Obtain cost multipliers from utility penalties, maintenance logs, or equipment manufacturers.
- **Energy Tariffs:** Use real-time tariffs to enhance accuracy for savings calculations.
