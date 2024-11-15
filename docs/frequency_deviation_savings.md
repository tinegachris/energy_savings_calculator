# Calculating Cost Savings from Frequency Stability

To calculate the cost savings for maintaining frequency stability at nominal using your time-series data, follow these steps:

## Step 1: Quantify Frequency Deviations

  1. **Determine Frequency Deviation (Δf):**

    For each timestamp in the data:

    - Δf = |f_measured - f_nominal|
    - f_measured: Supplied or load frequency (grid or corrected by BESS).
    - f_nominal: Nominal frequency (configurable e.g., 50 Hz or 60 Hz).

  2. **Categorize Frequency Deviations by Severity:**

    - Use industry standards or grid codes to define tolerable limits (e.g., ±0.1 Hz for most systems).
    - Define:
      - Acceptable deviations: |Δf| ≤ threshold.
      - Unacceptable deviations: |Δf| > threshold.

  3. **Frequency Stabilization Effect by BESS:**

    Measure improvement in frequency deviations:

    - Δf_reduction = |Δf_grid| - |Δf_BESS|

## Step 2: Assess the Impact of Frequency Deviations

  Frequency instability affects costs in several ways:

### A. Operational Inefficiencies

    1. **Energy Loss Due to Frequency Mismatch:**

      Deviations from nominal frequency affect the efficiency of rotating machines (e.g., motors, turbines, and generators).

      - Calculate mismatch energy:
        - E_Δf = P_load * time interval
      - Estimate energy efficiency loss using machine derating curves or performance data.

### B. Equipment Stress and Failure

    2. **Reduced Equipment Lifespan:**

      Frequency deviations increase mechanical stress on synchronous machines and sensitive electronics, leading to higher maintenance costs and shorter lifespans.

    3. **Cost of Unplanned Downtime:**

      Prolonged deviations can trigger protection systems, causing outages.

### C. Penalties from Utility Providers

    4. **Utility Charges for Grid Frequency Violations:**

      Some utilities impose penalties on consumers who exacerbate frequency instability (e.g., through overloading or exporting fluctuating power).

## Step 3: Quantify Cost Implications

### A. Equipment Derating Costs

  1. **Use machine derating curves (e.g., power loss per Hz deviation):**

    - Calculate lost power capacity:
      - P_loss = P_rated * Derating Factor(Δf)
    - Convert power loss into energy loss over time intervals:
      - E_loss = P_loss * time interval

  2. **Convert energy loss to cost:**

    - Cost_derating = E_loss * C_energy

### B. Maintenance and Replacement Costs

  1. **Estimate the increase in maintenance costs due to stress caused by grid frequency deviations (historical maintenance records can help):**

    - Cost_maintenance = (Failure Rate Increase) * (Repair or Replacement Cost)

  2. **Compare maintenance costs with and without BESS frequency correction.**

### C. Downtime Costs

  1. **Quantify outage costs due to prolonged frequency instability:**

    - Use the site's productivity rate or revenue loss per hour:
      - Downtime Cost = Downtime Hours * Revenue Loss Rate

  2. **Calculate avoided outages due to BESS corrections:**

    - Cost Savings_downtime = Downtime Cost_grid - Downtime Cost_BESS

## Step 4: Calculate Total Cost Savings

1. **Aggregate Savings from All Components:**

    - Total Cost Savings = Cost_derating + Cost_maintenance + Cost_downtime + Penalty Savings

2. **Compare savings over the analyzed period:**

    - Without BESS: Based on Δf_grid.
    - With BESS: Based on Δf_BESS.
