# Meter Tool for ATX Power v1.02

A Windows application for monitoring and calibrating a Nuvoton NADC24-based ATX power meter. It connects to the meter over a COM port and to Chroma instruments through NI-VISA.

🚀 Key Features
- Displays frequency, Vrms, Irms, power, and Vbulk with live charts.
- Reads device information and supports four-point calibration.
- Automates 5%–100% load sweeps and compares meter readings with Chroma references.
- Saves test results to timestamped output_yyyyMMdd_HHmmss.txt files.

📊 Hardware & Communication
Device	Interface	Function
NADC24 ATX power meter	USB-to-serial, 115200 8-N-1	Measurement and calibration
Chroma 66202 power meter	NI-VISA	Reference measurements
Chroma 63202A electronic load	NI-VISA	Programmable load
Calibration points: AC 100 V at 10% and 80% load; AC 230 V at 80% and 10% load (Points 1–4).

🛠️ Automated Load Test
Enter the DUT full-load current, settling delay, and repeat count, then click Start. The tool steps the load from 5% to 100% in 5% increments, reads voltage/current/power from both meters, calculates error, and logs the results.

💻 Environment & Build
- Windows, Visual Studio, and .NET Framework 4.7.2
- NI-VISA and instrument drivers for Chroma functions
Open Meter_Tool_for_ATX_Power_v1_02.sln, restore NuGet packages, and build the solution. Update PowerMeter66202_Address and DC_Load63202A_Address in Form1.cs to match the VISA addresses on your test PC.

⚠️ Note
The meter firmware and PC application must use compatible serial commands and auto-report format. Follow electrical safety procedures when testing power supplies and high-current loads.
