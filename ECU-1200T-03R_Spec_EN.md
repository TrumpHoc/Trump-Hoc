# TECHNICAL SPECIFICATION

**Product name:** Enclosure Cooling Unit (Cabinet Air Conditioner)
**Model:** ECU-1200T-03R
**Date of issue:** 2025/06/16
**Document No.:** TY-JS-250616-008
**Manufacturer:** Guangzhou Teyu Electromechanical Co., Ltd. (S&A)

> *The contents of this specification are for reference only. Guangzhou Teyu Electromechanical Co., Ltd. reserves the right of final interpretation and the right to change the contents without prior notice.*

---

## 1. Model Designation

```
ECU - 1200 - T - 03 - R
 │      │    │    │   └─ Function type: R – Standard; H – High-temperature;
 │      │    │    │                     EH – With electric heater; NCW – Condensate-free
 │      │    │    └───── Structure & mounting: 01 – Indoor type, internal side mounting;
 │      │    │                                 02 – Outdoor type, internal side mounting;
 │      │    │                                 03 – Indoor type, external side mounting
 │      │    └────────── Power supply: A – 1-phase 220 V/50 Hz; B – 1-phase 220 V/60 Hz;
 │      │                              T – 1-phase 220 V/50/60 Hz
 │      └─────────────── Cooling capacity (W)
 └────────────────────── Machine type: Enclosure Cooling Unit
```

*Note: This designation covers only the main product codes of the company and is not exhaustive. Please confirm the exact model with the company before ordering; the company reserves the right of final interpretation.*

---

## 2. Technical Data

### 2.1 Technical Parameters

| Item | Unit | ECU-1200T-03R |
|---|---|---|
| Power supply | – | AC 1-phase 220 V, 50/60 Hz |
| Ambient temperature range | °C | −5 ~ 50 |
| Cabinet temperature setting range | °C | 25 ~ 38 (factory setting 30 °C) |
| Rated cooling capacity (50/60 Hz) | W | 1200 / 1440 |
| Rated power input (50/60 Hz) | W | 680 / 760 |
| Rated current (50/60 Hz) | A | 3.0 / 3.6 |
| Refrigerant | – | R-134a |
| Refrigerant charge | g | 300 |
| Noise level | dB | ≤ 63 |
| Airflow – internal circuit | m³/h | 300 |
| Airflow – external circuit | m³/h | 500 |
| Power connection | – | Reserved wiring terminals |
| Power cord length | m | 2 |
| Net weight | kg | 28 |
| Gross weight | kg | 29 |
| Overall dimensions (L × W × H) | mm | 360 × 185 × 745 |
| Package dimensions (L × W × H) | mm | 430 × 255 × 820 |

**Notes:**
1. Cooling capacity is measured at ambient dry-bulb temperature 35 °C and cabinet internal dry-bulb temperature 35 °C (L35/L35).
2. For special operating conditions, please contact the company.

### 2.2 Main Components List

| No. | Component | Model | Brand |
|---|---|---|---|
| 1 | Compressor | 39F123N | Rechi (瑞智) |
| 2 | Evaporator | ZF-200×100 | S&A |
| 3 | Condenser | LN-240×200 | S&A |
| 4 | Evaporator fan | FLH175/042A | Xiangming (祥明) |
| 4 | Condenser fan | FLH190/045D | Xiangming (祥明) |
| 5 | Controller | STC-8000H | Elitech (精创) |

---

## 3. Refrigeration Schematic

Refrigerant circuit: **Compressor → Condenser (with condenser fan) → Filter drier → Capillary tube → Evaporator (with evaporator fan) → Compressor.**

---

## 4. Electrical Control

### 4.1 Electrical Schematic

- Supply: L, N, PE.
- Temperature controller (T °C) switches the line.
- M1 – Indoor (evaporator) fan, with run capacitor C1.
- M2 – Outdoor (condenser) fan, with run capacitor C2.
- M3 – Compressor, with run capacitor C3.

### 4.2 Wiring Diagram (Controller STC-8000H terminals)

| Terminal | Function |
|---|---|
| 1 | Compressor (output) |
| 2 | Common |
| 3 | Normally closed (NC) contact |
| 8 – 9 | Cabinet temperature sensor |
| 10 – 11 | Power supply 220 VAC |

Relay contact rating: 10 A / 277 VAC.

### 4.3 Electrical Cabinet Panel Cut-out Dimensions (mm)

| Item | Dimension |
|---|---|
| Mounting outline (W × H) | 350 × 755 |
| Return air opening (hot air, upper) | 300 (W) × 210 (H) |
| Supply air opening (cold air, lower) | 300 (W) × 95 (H) |
| Gap between return and supply openings | 85 |
| Fixing bolt holes | 6 × Ø8 |
| Bolt hole spacing – horizontal | 328 |
| Bolt hole spacing – vertical | 275 + 275 |
| Bolt hole offset from outline (top / side) | 10 / 30 (approx., refer to original drawing) |

---

## 5. Outline Drawing – Part Identification

| Location | Part |
|---|---|
| Front – top | Condenser air outlet (exhaust) |
| Front – middle | Condenser air inlet |
| Front – lower | Temperature controller |
| Rear (cabinet side) – upper | Return air inlet (hot air) |
| Rear (cabinet side) – lower | External air deflector |
| Rear (cabinet side) – lower | Supply air outlet (cold air) |

---
*Translated from the Chinese original (TY-JS-250616-008). In case of discrepancy, the Chinese original prevails.*
