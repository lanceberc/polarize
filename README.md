# polarize
Analyze sailboat performance data gathered from NEMA-0183 and NEMA-2000 (N2K) data sources

- Analyze sailboat race performance using data collected from NMEA bus
- Compute polar plots, per-leg strip charts, minute-by-minute summaries, and Expedition-ready text
- Generate reports in various formats - text files, graphics, and spreadsheets

NMEA-2000 .log files from Yacht Devices VDR Voyage Data Recorder are converted to
JSON by the canboat analyzer (https://www.github.com/canboat/canboat)

NMEA-0183 .nmea files from SEAiq are parsed directly

Configuration is in a per-regatta regatta.json file which defines the courses,
leg start/end times, and boat-specific parameters

Polars can be generated per-regatta or aggregated from multiple regattas

Requires Python3.x and additional modules:
- matplotlib for plotting (https://matplotlib.org)
- numpy for some statistics (https://numpy.org)
- scipy for Savitsky-Golay filtering (https://scipy.org)
- xlsxwriter to generate Excel-compatible .xlsx files (https://pypi.org/project/XlsxWriter)

![Strip Chart](Mayan_2019-09_RBBS_4_strip.png Strip Chart)
![Polar Chart](Mayan_aggregate_polars.png Polar Chart)
![Combined Polar Chart](Mayan_aggregate_combined_polars.png Combined Polar Chart)

