# Second-pass open-data reconstruction of the TPWRS IEEE-118 case

This run follows the companion Applied Energy paper more closely.

## What changed versus the first pass

- original IEEE-118 branch reactances scaled by 0.4
- 54 generator-side buses use the reported 27xTB6, 9xTC6, 5xTB10, 2xTC10, 11xTG3 fleet mix
- generator-side coupling uses the reported x'd templates as transformer-side coupling priors
- 100 MVA base, 50 Hz conversion and 1% load-damping initialization are explicit
- targeted open-data fitting retained around the north weak area and the 49/50 disturbance corridor

## Achieved headline numbers

- COI nadir: -0.0959 Hz
- Bus 48 initial RoCoF: -1.6933 Hz/s
- Bus 49 initial RoCoF: -1.8268 Hz/s
- Bus 08 nadir: -0.1170 Hz
- Bus 09 nadir: -0.1213 Hz
- Bus 48 nadir: -0.1578 Hz
- Bus 49 nadir: -0.1561 Hz
