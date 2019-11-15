# polarize
Analyze sailboat performance data gathered from NEMA-0183 and NEMA-2000 (N2K) data sources

## Synopsis

- Analyze sailboat race performance using data collected from NMEA bus
- Compute polar plots, per-leg strip charts, minute-by-minute summaries, and Expedition-ready text
- Generate reports in various formats - text files, graphics, and spreadsheets

It has become very easy and inexpensive to record raw yacht NMEA data, for instance:
- Directly via ~$250 Yacht Devices Voyage Data Recorder
- Indirectly via a B&G Zeus or Vulcan MFD and apps like SEAiq ($5 for iPhone/USA, $50 for Android/World)

polarize takes this data and converts it to more familiar forms - polar charts, strip charts, track files, spreadsheets, etc.

## Configuration

polarize expects two inputs:
1. A list of .json files describing the races and courses in a regatta (typically regatta.json)
2. A per-race NMEA data file pointed to from inside the regatta file

The regatta contains additional information including boat name, timezone offset, rudder correction, and COG/SOG sources.
Typically one creates a directory for each regatta since they share courses.

Polars can be generated per-regatta or aggregated from multiple regattas (polarize -polars */regatta.json)

## Output

polarize can generate several kinds of output:

Option | Effect
------ | ------
 \-strip | Create per-race strip graph files
 \-legs  | Create per-leg analysis
 \-minute | Create minute-by-minute reports in per-leg analysis
 \-spreadsheet | Create xlsx spreadsheet file
 \-polars | Create aggregate polar graph file
 \-exp | Create aggregate Expedition polar text file
 \-gpx | Create gpx track

## Notes
NMEA-0183 .nmea files from SEAiq are parsed directly

NMEA-2000 .log files from Yacht Devices VDR Voyage Data Recorder are converted to
JSON by the canboat analyzer (https://www.github.com/canboat/canboat)

Requires Python3.x and additional modules:
- matplotlib for plotting (https://matplotlib.org)
- numpy for some statistics (https://numpy.org)
- scipy for Savitsky-Golay filtering (https://scipy.org)
- xlsxwriter to generate Excel-compatible .xlsx files (https://pypi.org/project/XlsxWriter)

## Example Output

![Strip Chart](Mayan_2019-09_RBBS_4_strip.png)
![Polar Chart](Mayan_aggregate_polars.png)
![Combined Polar Chart](Mayan_aggregate_combined_polars.png)

