# polarize

Analyze sailboat performance data gathered from NMEA-0183 and NMEA-2000 (N2K) data sources

## Synopsis

- Analyze sailboat race performance using data collected from NMEA bus
- Compute polar plots, per-leg strip charts, minute-by-minute summaries, and Expedition-ready text
- Generate reports in various formats - text files, graphics, and spreadsheets

It has become very easy and relatively inexpensive to record raw NMEA data, for instance:
- Directly via ~$250 Yacht Devices Voyage Data Recorder
- Indirectly via N2K to WiFi gateways (such as any B&G Zeus or Vulcan MFD, Vesper AIS/WiFi gateways, and Yacht Devices YDWG NMEA-2000 dongles) and apps like SEAiq ($5 for iPhone/USA, $50 for Android/World)

polarize takes this data and converts it to more familiar forms - polar charts, strip charts, track files, spreadsheets, etc.

## Inputs/Configuration

polarize requires two inputs:
1. A list of .json files describing the races and courses in a regatta (typically regatta.json)
2. A per-race NMEA data file pointed to from inside the regatta file
   - .log files are assumed to have raw N2K data
   - .nmea files are assumed to have NMEA-0183 text sentences

The regatta file contains additional information including boat name, timezone offset, rudder correction, and COG/SOG sources.
Typically one creates a directory for each regatta since they share courses.

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
 
 Polars can be generated per-regatta or aggregated from multiple regattas (polarize -polars \*/regatta.json)

## Notes

### Input data format
NMEA-0183 .nmea files (i.e. recorded with SEAiq) are parsed directly

NMEA-2000 .log files from Yacht Devices VDR Voyage Data Recorder are converted to
JSON by the canboat analyzer (https://www.github.com/canboat/canboat)

### Conflicting NMEA sources
It's becoming common for boats to have multiple sources of similar data. For instance COG/SOG data
on a boat with a B&G MFD and ZG100 GPS/heading sensor will disagree due to different damping heuristics,
sometimes by quite a bit. The N2K Source ID can be specified in the regatta file for N2K data to choose
which device to use.
If naively converted to
NMEA-0183 by a WiFi gateway the source ID field is lost making it hard/impossible to filter for only one source.
This may confuse polarize.

### Software Environment
polarize requires Python3.x and additional modules:
- matplotlib for plotting (https://matplotlib.org)
- numpy for some statistics (https://numpy.org)
- scipy for Savitsky-Golay filtering (https://scipy.org)
- xlsxwriter to generate Excel-compatible .xlsx files (https://pypi.org/project/XlsxWriter)

### Localization
polarize requires kludgey internal configuration:
- The ANALYZE variable must point to the local copy of analyze to enable N2K data conversion
- Polar wind ranges are set manually in a table. The default is good for a particular schooner in typical San Francisco Bay wind ranges.

## Example Output

![Strip Chart](Mayan_2019-09_RBBS_4_strip.png)
![Polar Chart](Mayan_aggregate_polars.png)
![Combined Polar Chart](Mayan_aggregate_combined_polars.png)

## License
Copyright 2019 Lance Berc

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
