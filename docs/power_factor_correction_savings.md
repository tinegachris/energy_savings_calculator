# Calculating Cost Savings from Power Factor Correction

To calculate cost savings due to power factor (PF) correction from your solar PV and BESS system using real power (P) and apparent power (S) time series, follow these steps:

## Step 1: Quantify the Power Factor Improvement

1. **Calculate Power Factor for Each Timestamp:**

```plaintext
PF_measured = P / S
```

- P: Real power (kW).
- S: Apparent power (kVA).

2. **Categorize Power Factor by Tolerance:**

- Compare PF_measured against a target PF (configurable e.g., 0.90 or 0.95, as required by utilities).
- Define:
  - Unacceptable PF: PF_measured < PF_target.
  - Acceptable PF: PF_measured >= PF_target.

3. **Calculate Corrected PF:**

- Incorporate the reactive power contribution of the solar PV and BESS to estimate the corrected power factor:

  ```plaintext
  PF_corrected = P / sqrt(P^2 + (Q_corrected)^2)
  ```

- Q_corrected = Q_measured - Q_BESS+PV, where:

  - Q_measured = sqrt(S^2 - P^2): Reactive power demand before correction.
  - Q_BESS+PV: Reactive power supplied by the system.

4. **Quantify Improvement:**

  ```plaintext
  Delta PF = PF_corrected - PF_measured
  ```

## Step 2: Assess the Impact of Power Factor

Reactive power costs are often linked to utility charges and system inefficiencies.

### A. Utility Charges for Low Power Factor

1. **Many utilities impose penalties for low PF:**

- Reactive energy cost (kVARh).
- Direct penalties for PF below a threshold.

2. **Calculate Reactive Energy (kVARh):**

- For each timestamp:

  ```plaintext
  Q = sqrt(S^2 - P^2)
  ```

- Sum up reactive energy over the month:

  ```plaintext
  E_reactive = sum Q * interval duration (hours)
  ```

3. **Reactive Power Costs:**

- Before correction:

  ```plaintext

  Cost_reactive, before = E_reactive, measured * C_kVARh
  ```

- After correction:

  ```plaintext
  Cost_reactive, after = E_reactive, corrected * C_kVARh
  ```

4. **Savings:**

  ```plaintext
  Cost Savings = Cost_reactive, before - Cost_reactive, after
  ```

### B. Equipment Efficiency Gains

1. **Lower Reactive Power Demand Reduces Transmission Losses:**

- Transmission and distribution losses scale with reactive power. Estimate loss reduction:

  ```plaintext
  Delta Losses = k * (Q_before^2 - Q_after^2)
  ```

  - k: System loss factor (can be obtained from utility or system data).

2. **Convert loss reduction into energy savings:**

  ```plaintext
  Energy Savings = Delta Losses * interval duration
  ```

3. **Associate savings with cost:**

  ```plaintext
  Cost Savings_losses = Energy Savings * C_energy
  ```

### C. Improved Equipment Lifespan

1. **Reduced Equipment Stress:**

- Motors and transformers operate more efficiently with higher PF.
- Quantify lifespan increase using derating curves or maintenance data.

2. **Cost Reduction:**

  ```plaintext
  Cost Savings_maintenance = Reduction in Repairs + Avoided Replacements
  ```

## Step 3: Account for Avoided Penalties

1. **Determine Utility Penalty Thresholds:**

- For each period where PF_measured < PF_penalty, calculate the penalty:

  ```plaintext
  Penalty_before = E_reactive, excess * C_penalty
    ```

2. **Post-Correction Penalty Avoidance:**

  ```plaintext
  Penalty Savings = Penalty_before - Penalty_after
  ```

## Step 4: Total Cost Savings

1. **Aggregate savings over the month:**

  ```plaintext
  Total Cost Savings = Cost Savings_reactive + Cost Savings_losses + Cost Savings_maintenance + Penalty Savings
  ```

2. **Compare scenarios:**

- Without correction: Based on PF_measured.
- With correction: Based on PF_corrected.
