# Harmonic Distortion Savings Calculation

To calculate **Harmonic Distortion (HD)** and estimate the cost savings due to its mitigation using the data provided, follow these steps:

### **1. Understanding Total Harmonic Distortion (THD)**

- **Current THD (\(THD_I\))**: A measure of how much the current deviates from a pure sinusoidal waveform.
- **Voltage THD (\(THD_V\))**: A measure of voltage waveform distortion.
- Both are expressed as percentages:
  \[
  THD = \frac{\sqrt{\sum_{n=2}^\infty (H_n)^2}}{H_1} \times 100\%
  \]
  - \(H_1\): Fundamental component.
  - \(H_n\): \(n\)-th harmonic.

---

### **2. Input Data**

- **Time Series Data**:
  - Timestamp.
  - \(THD_I\): Total harmonic distortion for current.
  - \(THD_V\): Total harmonic distortion for voltage.

- Collected every 5 minutes for a month.

---

### **3. Approach to Harmonic Cost Calculation**

#### **A. Identify Harmonic Standards and Limits**

- Identify acceptable THD levels based on standards like IEEE 519 or local regulations:
  - For voltage: \(THD_V \leq 5\%\) (for systems below 69 kV).
  - For current: Acceptable limits depend on load type and size.

- Separate time-series data into:
  - **Compliant periods**: Where \(THD_V\) and \(THD_I\) are below the threshold.
  - **Non-compliant periods**: Where \(THD_V\) or \(THD_I\) exceed the threshold.

#### **B. Quantify Non-Compliant Harmonic Energy**

1. **Calculate Total Apparent Energy (\(E_S\)) for Each Period**:
   - Using EMS data:
     \[
     E_S = S \cdot \text{Interval (hours)}
     \]
     - \(S\): Apparent power (kVA).

2. **Attribute Energy to Non-Compliant Periods**:
   - If \(THD_V\) or \(THD_I\) exceeds the threshold, record the corresponding \(E_S\) as non-compliant energy.

#### **C. Link to Cost Factors**

1. **Energy Losses**:
   - Harmonics increase losses in transformers and conductors.
   - Estimate loss factor (\(k\)) for harmonics:
     \[
     \Delta E_{\text{loss}} = k \cdot E_S \cdot \frac{THD}{100}
     \]
     - Typical \(k\) values can be derived from empirical data or transformer specs.

2. **Equipment Aging**:
   - Harmonics accelerate wear on components:
     - Motors.
     - Transformers.
     - Capacitors.
   - Approximate replacement cost:
     \[
     \text{Cost}_{\text{aging}} = \text{Annual Maintenance} \cdot \frac{\Delta \text{Wear Rate}}{\text{Normal Wear Rate}}
     \]

3. **Utility Penalties**:
   - Some utilities penalize customers for exceeding harmonic limits:
     \[
     \text{Penalty} = \text{Excess Harmonic Energy (kVAh)} \cdot \text{Penalty Rate}
     \]

#### **D. Harmonic Mitigation Savings**

1. **Reduction in Energy Losses**:
   - Compare harmonic energy before and after mitigation.
   - Calculate savings:
     \[
     \text{Savings}_{\text{loss}} = \Delta E_{\text{loss (before)}} - \Delta E_{\text{loss (after)}}
     \]

2. **Avoided Penalties**:
   - Quantify penalties avoided by maintaining compliance.

3. **Improved Equipment Lifespan**:
   - Estimate maintenance and replacement cost savings:
     \[
     \text{Savings}_{\text{aging}} = \text{Cost}_{\text{aging(before)}} - \text{Cost}_{\text{aging(after)}}
     \]
