#!/usr/local/bin/python3
# -*- coding: utf-8 -*-

"""
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
"""

# Analyze sailboat race performance using data collected from NMEA bus
# Compute polar plots, per-leg strip charts, minute-by-minute summaries, and Expedition-ready text
# Generate reports in various formats - text files, graphics, and spreadsheets

# NMEA-2000 .log files from Yacht Devices VDR Voyage Data Recorder are converted to
# JSON by the canboat analyzer (https://www.github.com/canboat/canboat)

# NMEA-0183 .nmea files from SEAiq are parsed directly

# Configuration is in a per-regatta regatta.json file which defines the courses,
# leg start/end times, and boat-specific parameters

# Polars can be generated per-regatta or aggregated from multiple regattas

# Requires Python3.x and additional modules:
#   matplotlib for plotting                             https://matplotlib.org/
#   numpy for some statistics                           https://numpy.org
#   scipy for Savitsky-Golay filtering                  https://scipy.org
#   xlsxwriter to generate Excel-compatible .xlsx files https://pypi.org/project/XlsxWriter/

deg = u'\N{DEGREE SIGN}'

from dataclasses import dataclass
import os.path
import datetime
import subprocess
from math import sqrt, sin, cos, pi, tau, radians, degrees, atan2, fmod
from statistics import mean
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import scipy.signal
import json
import argparse
import xlsxwriter

regattalist = []
regattas = {}

rudderCorrection = 0.0

# Build the N2K analyzer from www.github.com/canboat/canboat
# It's written in C. Pretty straight forward.
# 'analyzer' converts raw binary N2K to somewhat verbose JSON, but it's a small price to pay
ANALYZER = "/Users/lance/src/canboat/canboat/rel/darwin-x86_64/analyzer"

tzoffset = None
sampleSeconds = 10
reportSeconds = 60

# YYYY-MM-DD HH:MM:SS.sss
boards = { 'Port': { }, 'Stbd': { } }

raceRawFields = ['AWA', 'AWS', 'STW', 'RUD', 'COG', 'SOG', 'HDG', 'LATLON', 'TWA', 'TWD', 'TWS']
legRawFields = ['AWA', 'AWS', 'STW', 'RUD', 'COG', 'SOG', 'HDG', 'LATLON', 'TWA', 'TWD', 'TWS']
legFields = ['Time', 'TWA', 'TWD', 'TWS', 'AWA', 'AWS', 'STW', 'RUD', 'COG', 'SOG', 'HDG']

# Min and max true wind angles for polars - assume if it's outside the range we're tacking or gybing
minTWA = 25.0
maxTWA = 165.0

# Min and max apparent wind angles for a leg - assume if it's outside the range we're tacking or gybing
minAWA = 25.0
maxAWA = 165.0

# Set latlonSource and cogsogSource to 1 for B&G or 21 for the Vesper XB8000 - may vary per-boat
LATLONSOURCE = None
COGSOGSOURCE = None

# Color map for most lines in strip and polar plots. Chosen to look 'nice'
#cmap = plt.cm.Dark2.colors
# Dark Pastels from https://www.schemecolor.com/ with yellow moved to end
cmap = [ '#C23B23', '#F39A27', '#03C03C', '#579ABE', '#976ED7', '#EADA52']

# Convert meters per second to knots
def ms2kts(ms):
    return(ms * 1.94384)

def addtz(ts):
    return("%s" % (ts + tzoffset))

def parse_regatta(fn):
    # Parse a regatta description, which is a JSON file composed of a list of dictionary elements.
    # Each element is a "regatta", "race", or "course"

    with open(fn, "r") as f:
        try:
            j = json.load(f)
        except:
            print("Couldn't load %s as JSON" % (fn))
            raise

    r = None
    for e in j:
        # Look at each element; so far "regatta", "race", and "course" are defined
        if "regatta" in e:
            name = e["regatta"]
            regattalist.append(name)
            regattas[name] = e # copy all params into regatta
            regatta = regattas[name]
            p, n = os.path.split(fn)
            regatta["path"] = p + '/' if p != "" else "./"
            regatta["races"] = []
            regatta["courses"] = {}
            global rudderCorrection
            rudderCorrection = 0.0 if not 'rudderCorrection' in e else float(e['rudderCorrection'])
            print("## Setting rudder correction to %4.1f%s" % (rudderCorrection, deg))
            global LATLONSOURCE, COGSOGSOURCE
            LATLONSOURCE = None if not 'latlonSource' in e else int(e['latlonSource'])
            COGSOGSOURCE = None if not 'cogsogSource' in e else int(e['cogsogSource'])
            tzoffset = datetime.timedelta() if not 'tz' in e else datetime.timedelta(hours=int(e['tz']))
            print("## Parse Regatta %s" % (name))
        elif "race" in e:
            ## print("## Parse Race %r" % (e))
            e['startts'] = datetime.datetime.strptime(e['start'], '%Y-%m-%dT%H:%M:%S') - tzoffset
            e['endts'] = datetime.datetime.strptime(e['end'], '%Y-%m-%dT%H:%M:%S') - tzoffset
            for f in raceRawFields:
                e[f] = []
            regatta["races"].append(e) # Add this to the list of races
            print("## Parse Race %s - Course %s %s - %s" % (e['race'], e['course'], e['startts'], e['endts']))

        elif "course" in e:
            regatta['courses'][e["course"]] = e # Add this to the dictionary of courses
            print("## Parse Course %s" % (e['course']))
        else:
            print("Unknown regatta element '%s'" % (e))

# https://www.eye4software.com/hydromagic/documentation/nmea0183/
# https://www.trimble.com/OEM_ReceiverHelp/V4.44/en/NMEA-0183messages_MessageOverview.html
# https://www.gpsinformation.org/dale/nmea.htm#intro
"""
Sentences seen at Doc Brown's Lab - B&G Zeus MFD, ZG100, SDT transducer, H5000 masthead, Off-brand heading sensor
Recorded with SeaIQ via a Yacht Devices YDG wireless gateway
Count Sntnce *     Description
57854 $SDHDG * HDG Heading w/ magnetic variation
17667 $GPGSV   GSV Satellites in view
11774 $WIMWV * MVW Wind speed and angle - Two sentences, relative and true?
9197  $PSIQREC * SeaIQ recording time stamps
5889  $SDDPT   DPT Depth
5889  $IIXDR * XDR Transducer measurement - includes rudder angle?
5889  $GPXTE   XTE Cross-track error
5889  $GPRMB   RMB Navigation info
5889  $GPGSA   GSA GPS Info
5889  $GPGLL * GLL GPS-based lat/lon
5889  $GPGLC   GLC Loran-based lat/lon
5889  $GPGGA   GGA GPS fix data
5889  $GPBWR   BWR Bearing and distance to waypoint - Ruhmb line
5889  $GPBWC   BWC Bearing and distance to waypoint - Great circle
5888  $WIMWD * MWD Wind Direction and Speed, with respect to north
5888  $SDVLW   VLW Distance through water
5888  $SDVHW * VHW Water speed and heading
5888  $SDMTW   MTW Water temp
5888  $GPZDA * ZDA Date & Time (UTC)
5888  $GPVTG * VTG Track made good & ground speed (cog/sog?)
5888  $GPBOD   BOD Bearing - waypoint to waypoint
5888  $GPAPB   APB Autopilot 'B'
5888  $GPAAM   AAM Waypoint arrival alarm
5887  $GPRMC   RMC Navigation info
"""

# NMEA 0183 sentences we use
interesting_sentences = ["$SDHDG", # Heading
                         "$WIMWV", # Wind data
                         "$PSIQREC", # SeaIQ Time stamps
                         "$IIXDR", # Transducer measurement w/ rudder angle
                         "$GPGLL", # LAT/LON
                         "$SDVHW", # Water speed - and heading?
                         "$GPZDA", # Time stamp
                         "$GPVTG"] # COG/SOG
    
def legInit(l):
    for f in legRawFields:
        l[f] = []
    l['startts'] = datetime.datetime.strptime(l['start'], '%Y-%m-%dT%H:%M:%S') - tzoffset
    l['endts'] = datetime.datetime.strptime(l['end'], '%Y-%m-%dT%H:%M:%S') - tzoffset
    l['duration'] = l['endts'] - l['startts']
    l['sindex'] = {}
    l['eindex'] = {}
    print("## Init leg %s - %s" % (l['startts'], l['endts']))

interesting_sentences = ["$SDHDG", # Heading
                         "$WIMWV", # Wind data
                         "$PSIQREC", # SeaIQ Time stamps
                         "$GPGLL", # LAT/LON
                         "$SDVHW", # Water speed - and heading?
                         "$GPZDA", # Time stamp
                         "$GPVTG"] # COG/SOG

def parse_race_0183(regatta, r):
    with open(regatta['path'] + r['data'], 'r') as f:
        r['variation'] = 0
        r['LATLON'] = []
        variation = 0
        ts = datetime.datetime(year=1970, day=1, month=1, hour=0, minute=0, second=0)
        sampleCount = 0
        sentenceCount = 0

        while True:
            line = f.readline()
            if not line:
                break
            fields = line.split(',')

            # I'm tempted to select on the entire field, but I don't know if some brands emit
            # some sentences with different talker IDs
            talker = fields[0][:-3]
            sentence = fields[0][-3:]

            # Look for timestamps first.
            if sentence == "REC":
                # SeaIQ recording timestamp
                # $PSIQREC,0,1,1572722175.566,20191102,121615
                # print("## REC %s: %s %s %s" % (line, fields[3], fields[4], fields[5]))
                y = int(fields[4][0:4])
                m = int(fields[4][4:6])
                d = int(fields[4][6:8])
                h = int(fields[5][0:2])
                minute = int(fields[5][2:4])
                s =  int(fields[5][4:6])
                ms = int(fields[3][-3:])
                tslocal = datetime.datetime(year=y, month=m, day=d, hour=h, minute=minute, second=s, microsecond=ms*1000)
                newts = datetime.datetime.utcfromtimestamp(float(fields[3]))
                delta = newts - ts
                # Ignore the SeaIQ timestamps - use the ZDA sentence from the B&G
                #print("## New time REC %s delta %s" % (newts.strftime('%Y-%m-%dT%H:%M:%S'), delta))
                #ts = newts

            elif sentence == "ZDA":
                # GPS timestamp
                # $GPZDA,191614,02,11,2019,07,00*4D
                # $GPZDA,UTC,Day,Month,Year,TZ-hours,TZ-minutes
                # UTC is in HHMMSS.xxxx
                y = int(fields[4])
                m = int(fields[3])
                d = int(fields[2])
                h = int(fields[1][0:2])
                minute = int(fields[1][2:4])
                s =  int(fields[1][4:6])
                newts = datetime.datetime(year=y, month=m, day=d, hour=h, minute=minute, second=s)
                delta = newts - ts
                #print("## New time ZDA %s delta %s" % (newts.strftime('%Y-%m-%dT%H:%M:%S'), delta))
                ts = newts

            if (ts < r['startts']):
                # Not yet in the race
                # print("## Not leg %s vs %s" % (ts.strftime('%Y-%m-%dT%H:%M:%S'), r['startts'].strftime('%Y-%m-%dT%H:%M:%S')))
                continue

            if (ts >= r['endts']):
                print("## Done %s - %s %d of %d samples" % (r['startts'], r['endts'], sampleCount, sentenceCount))
                break

            sentenceCount += 1
            if sentence == "GLL":
                # LAT/LON
                # $GPGLL,3748.8071,N,12227.9801,W,191614,A,A*5D
                # Remember to make South and West negative
                lat1, lat2 = fields[1].split('.')
                n = fields[2] == 'N'
                lat = (float(lat1[0:-2]) + (float(lat1[-2:] + '.' + lat2)/60)) * (1 if n else -1)

                lon1, lon2 = fields[3].split('.')
                e = fields[4] == 'E'
                lon = (float(lon1[0:-2]) + (float(lon1[-2:] + '.' + lon2)/60)) * (1 if e else -1)

                r['LATLON'].append((ts, lat, lon))
                sampleCount += 1

            elif sentence == "HDG":
                # Heading
                # $SDHDG,336.4,,,13.3,E*06
                if fields[4] != '':
                    east = fields[5] == 'E'
                    variation = float(fields[4]) * (1 if east else -1)
                    if r['variation'] != variation:
                        r['variation'] = variation
                        variation = variation
                        # print("Setting compass variation to %.1f" % (variation))
                hdg = float(fields[1])
                r['HDG'].append((ts, hdg))
                sampleCount += 1

            elif sentence == "MWV":
                # Wind data
                # $WIMWV,336.8,R,18.7,N,A*13
                # $WIMWV,324.6,T,13.0,N,A*14
                if fields[4] == 'N':
                    speed = float(fields[3])
                elif fields[4] == 'M':
                    speed = m2k(float(fields[3]))
                if fields[2] == 'R':
                    # Apparent wind - should make sure it's in knots
                    awa = float(fields[1])
                    awa = awa if awa < 180 else awa - 360
                    r['AWA'].append((ts, awa))
                    r['AWS'].append((ts, speed))
                if fields[2] == 'T':
                    # True wind - should make sure it's in knots
                    #l['TWA'].append((ts, float(fields[1])))
                    #l['TWS'].append((ts, speed))
                    #sampleCount += 1
                    #  We compute this from smoothed data rather than trust the instrument's calc
                    pass

            elif sentence == "MWD":
                # Wind Direction and Speed, with respect to north
                # $WIMWD,70.4,T,57.1,M,4.7,N,2.4,M*5F
                #twd = float(fields[3])
                #tws = float(fields[5])
                # Not saving these for now - not sure I believe them
                pass

            elif sentence == "VHW":
                # Speed through Water (with heading?)
                # $SDVHW,348.7,T,335.4,M,5.7,N,10.6,K*7E
                #hdg = float(fields[3]) # Do not believe heading from boat speed sensor
                stw = float(fields[5])
                r['STW'].append((ts, stw))
                sampleCount += 1
            elif sentence == "VTG":
                # COG/SOG
                # $GPVTG,342.9,T,329.5,M,5.4,N,10.1,K,A*13
                cog = float(fields[3]) # use magnetic COG
                sog = float(fields[5])
                r['COG'].append((ts, cog))
                r['SOG'].append((ts, sog))
                sampleCount += 1
    print("## Done Race %s %s - %s kept %d of %d pgns" % (r['race'], r['startts'], r['endts'], sampleCount, sentenceCount))

            
# The N2k PGNs we need for sailing performance. Ignores things like routes/waypoints, AIS, etc.
interesting_pgns = [127245, # Rudder angle
                    127250, # Vessel heading
                    127258, # Magnetic variation
                    128259, # Speed through water (boatspeed)
                    129025, # Position (lat/lon)
                    129026, # COG / SOG
                    130306] # Wind data

def parse_race_n2k(regatta, r):
    # Parse the NMEA data for a single race.
    # Data starts in a .log file from the Yacht Devices YDVRConv app
    # Use the analyzer to convert it to JSON, filtering for the interesting PGNs

    c = regatta["courses"][r['course']]
    print("## Parse Race %s N2K" % (r['race']))
    jsonfile = regatta['path'] + regatta["basefn"] + '_' + r['race'] + ".json"

    # If we don't have a JSON file, run the raw log through the analyzer filtering for the PGNs we care about
    if not os.path.isfile(jsonfile):
        pgnpat = '(' + str(interesting_pgns[0])
        for pgn in interesting_pgns[1:]:
            pgnpat += '|' + str(pgn)
        pgnpat += ')'

        logfile = regatta["path"] + r['data']
        print("Converting N2K %s to %s" % (logfile, jsonfile))

        with open(logfile, 'r') as infile, open(jsonfile, 'wb') as outfile:
            p1 = subprocess.Popen([ ANALYZER, '-json' ], stdin=infile, stdout=subprocess.PIPE)
            p2 = subprocess.Popen([ '/usr/bin/grep', '-E', pgnpat], stdin=p1.stdout, stdout=outfile)
            retcode = p2.wait()
            if retcode == None:
                print("N2K analyze pipe not terminated?")
            elif retcode != 0:
                print("N2K analyze pipe exited with code %d" % (retcode))
        
    with open(jsonfile, "r") as f:
        r['variation'] = 0
        r['LATLON'] = []
        variation = 0
        sampleCount = 0
        pgnCount = 0

        while True:
            line = f.readline()
            if not line:
                break
            j = json.loads(line)
            #print("Line %s" % (line))
            #print("JSON %r" % (j))
            ts = datetime.datetime.strptime(j['timestamp'][:-4], '%Y-%m-%d-%H:%M:%S')
            pgn = j['pgn']

            if pgn == 127258:
                #{"timestamp":"2019-10-20-19:04:56.535","prio":7,"src":1,"dst":255,"pgn":127258,"description":"Magnetic Variation","fields":{"SID":53,"Variation":13.3}}
                variation = j['fields']['Variation']
                if variation != r['variation']:
                    r['variation'] = variation
                    # print("## Setting compass variation to %f" % (variation))

            if (ts < r['startts']):
                # Not yet on the new leg
                continue

            # print("%s vs %s" % (time, r["legs"][leg]["start"]))
            if (ts >= r['endts']):
                print("## Done Race %s %s - %s kept %d of %d pgns" % (r['race'], r['startts'], r['endts'], sampleCount, pgnCount))
                break

            pgnCount += 1
            if pgn == 129025:
                # Record LAT/LON even when not on a leg
                #{"timestamp":"2019-10-20-19:04:57.588","prio":2,"src":1,"dst":255,"pgn":129025,"description":"Position, Rapid Update","fields":{"Latitude":37.8489280,"Longitude":-122.4480448}}
                #{"timestamp":"2019-10-20-19:04:57.598","prio":2,"src":21,"dst":255,"pgn":129025,"description":"Position, Rapid Update","fields":{"Latitude":37.8489088,"Longitude":-122.4480384}}
                if (LATLONSOURCE == None) or (j['src'] == LATLONSOURCE):
                    r['LATLON'].append((ts, j['fields']['Latitude'], j['fields']['Longitude']))
                    sampleCount += 1
            
            elif pgn == 127245:
                #{"timestamp":"2019-10-20-19:04:56.206","prio":2,"src":204,"dst":255,"pgn":127245,"description":"Rudder","fields":{"Instance":252,"Direction Order":0}}
                #{"timestamp":"2019-10-20-19:04:56.206","prio":2,"src":204,"dst":255,"pgn":127245,"description":"Rudder","fields":{"Instance":0,"Position":19.2}}
                if j['fields']['Instance'] == 0:
                    rudder = j['fields']['Position'] + rudderCorrection
                    r['RUD'].append((ts, rudder))
                sampleCount += 1

            elif pgn == 127250:
                #{"timestamp":"2019-10-20-19:04:56.308","prio":2,"src":204,"dst":255,"pgn":127250,"description":"Vessel Heading","fields":{"Heading":280.1,"Reference":"Magnetic"}}
                heading = j['fields']['Heading']
                r['HDG'].append((ts, heading))
                sampleCount += 1

            elif pgn == 128259:
                #{"timestamp":"2019-10-20-19:04:56.538","prio":2,"src":35,"dst":255,"pgn":128259,"description":"Speed","fields":{"SID":6,"Speed Water Referenced":1.35,"Speed Water Referenced Type":"Paddle wheel"}}
                ms = j['fields']['Speed Water Referenced']
                stw = ms2kts(ms)
                r['STW'].append((ts, stw))
                sampleCount += 1

            elif pgn == 129026:
                #{"timestamp":"2019-10-20-19:04:57.891","prio":2,"src":1,"dst":255,"pgn":129026,"description":"COG & SOG, Rapid Update","fields":{"SID":54,"COG Reference":"True","COG":281.5,"SOG":1.38}}
                #{"timestamp":"2019-10-20-19:04:58.001","prio":2,"src":21,"dst":255,"pgn":129026,"description":"COG & SOG, Rapid Update","fields":{"COG Reference":"True","COG":187.6,"SOG":1.47}}
                if (COGSOGSOURCE == None) or (j['src'] == COGSOGSOURCE):
                    cog = j['fields']['COG'] if j['fields']['COG Reference'] != "True" else j['fields']['COG'] - variation
                    sog = ms2kts(j['fields']['SOG'])
                    r['COG'].append((ts, cog))
                    r['SOG'].append((ts, sog))
                    sampleCount += 1

            elif pgn == 130306:
                #{"timestamp":"2019-10-20-19:04:58.009","prio":2,"src":9,"dst":255,"pgn":130306,"description":"Wind Data","fields":{"SID":0,"Wind Speed":7.18,"Wind Angle":281.8,"Reference":"Apparent"}}
                aws = ms2kts(j['fields']['Wind Speed'])
                awa = j['fields']['Wind Angle'] 
                awa = awa if awa <= 180.0 else awa - 360.0
                r['AWS'].append((ts, aws))
                r['AWA'].append((ts, awa))
                sampleCount += 1

def parse_race(regatta, r):
    if r['data'][-5:] == '.nmea':
        parse_race_0183(regatta, r)
    elif r['data'][-4:] == '.log':
        parse_race_n2k(regatta, r)
    else:
        print("Unknown NMEA format file %s" % (r['data']))

def analyze_race(regatta, r):
    # Find the start and end indices for each parameter on each leg
    # This is a few more lines than using a list comprehension, but it's O(n) instead of O(n^2)
    for l in r['legs']:
        legInit(l)

    for field in legFields:
        if field == 'Time':
            continue
        i = 0
        for leg, l in enumerate(r['legs']):
            if not field in r:
                # Don't have this field in the raw data
                l['sindex'][field] = 0
                l['eindex'][field] = 0
                print("## Leg %d No such field %s" % (leg, field))
                break
            while (i < len(r[field])) and (r[field][i][0] < l['startts']):
                i += 1
            l['sindex'][field] = i
            while (i < len(r[field])) and (r[field][i][0] < l['endts']):
                i += 1
            # Last data item for this leg
            l['eindex'][field] = i
            #print("## Leg %d Field %s[%d:%d]" % (leg, f, l['sindex'][f], l['eindex'][f]))

def analyze_leg(regatta, r, leg):
    # Coalesce the data from each leg into a list of 10 second samples
    # Compute the TWS, TWD, and TWA for each sample
    c = regatta['courses'][r['course']]
    l = r['legs'][leg]

    print("## Analyze %s Race %s Leg %d %s (%.2fnm @ %03d%s) - (%s - %s) %s"
          % (regatta['regatta'], r['race'], leg+1, c["legs"][leg]["label"], c["legs"][leg]["distance"], c["legs"][leg]["bearing"], deg, l['start'], l['end'], l['duration']))

    # Chop the leg into 10 second buckets
    # Look for tacks and gybes
    bucketDelta = datetime.timedelta(seconds=sampleSeconds)
    bucketStart = l['startts']

    index = {}
    for field in ['AWA', 'AWS', 'STW', 'SOG', 'RUD', 'TWS', 'COG', 'HDG', 'TWD']:
        index[field] = l['sindex'][field]
    
    l['samples'] = []
    while bucketStart < l['endts']:
        bucket = {}
        bucketEnd = bucketStart + bucketDelta
        bucketEnd = min(bucketEnd, l['endts'])
        bucket['ts'] = bucketStart

        for field in ['AWA', 'AWS', 'STW', 'SOG', 'RUD', 'TWS']:
            total = 0.0
            count = 0
            # Advance to the beginning of the run of data for this bucket
            while index[field] < l['eindex'][field] and r[field][index[field]][0] < bucketStart:
                index[field] += 1
            # Run through the valid data summing the value
            while index[field] < l['eindex'][field] and r[field][index[field]][0] < bucketEnd:
                d = r[field][index[field]]
                #if len(d) < 2:
                #    print("Field %s data element too small: %r" % (field, d))
                total += d[1]
                count += 1
                index[field] += 1
            # Take the average for the bucket
            bucket[field] = None if count == 0 else total / count
        
        markBearing = c["legs"][leg]["bearing"]
        # If we're going north-ish, convert COG and Heading ranges to [-180, 180] in case some samples straddle due north
        # This is to catch the "averaging samples around due north degress leads to due south" problem - mean([0,359.999]) = 180
        for field in ['COG', 'HDG', 'TWD']:
            total = 0.0
            count = 0
            # Advance to the beginning of the run of data for this bucket
            while index[field] < l['eindex'][field] and r[field][index[field]][0] < bucketStart:
                index[field] += 1
            while index[field] < l['eindex'][field] and r[field][index[field]][0] < bucketEnd:
                d = r[field][index[field]]
                # If we're going north-ish, convert COG and Heading ranges to [-180, 180] in case some samples straddle due north
                total += d[1] if (markBearing > 90 and markBearing < 270) or (d[1] <= 180) else d[1] - 360.0
                count += 1
                index[field] += 1
            if count == 0:
                bucket[field] = None
            else:
                bucket[field] = total / count
                bucket[field] += 0.0 if (markBearing > 90 and markBearing < 270) or (bucket[field] > 0) else 360.0 # Convert back to 0 - 360 if needed
        
        """
        Compute TWS and TWD for this bucket

        AWA = + for Starboard, – for Port
        AWD = H + AWA ( 0 < AWD < 360 )
        u = SOG * Sin (COG) – AWS * Sin (AWD)
        v = SOG * Cos (COG) – AWS * Cos (AWD)
        TWS = SQRT ( u*u + v*v )
        TWD = ATAN ( u / v )
        """

        if bucket['HDG'] == None or bucket['COG'] == None or bucket['SOG'] == None or bucket['AWA'] == None or bucket['AWS'] == None:
            # Can't compute True Wind w/o HDG, COG, SOG and AWA, AWS
            bucket['TWD'] = None
            bucket['TWS'] = None
            bucket['TWA'] = None
        else:
            hdg = radians(bucket['HDG'])
            aws = bucket['AWS']
            awa = radians(bucket['AWA'])
            cog = radians(bucket['COG'])
            sog = bucket['SOG']

            # Compute TWD, TWS if it's not done by the instruments
            if bucket['TWD'] == None:
                awd = fmod(hdg + awa, tau) # Compensate for boat's heading
                u = (sog * cos(cog)) - (aws * cos(awd))
                v = (sog * sin(cog)) - (aws * sin(awd))
                tws = sqrt((u*u) + (v*v))
                # Now we want to know where it's from, not where it's going, so add pi
                twd = degrees(fmod(atan2(v,u)+pi, tau))

                bucket['TWS'] = tws
                bucket['TWD'] = twd

            # Compute TWA from TWD and Heading
            twa = fmod((bucket['TWD'] - bucket['HDG']) + 360.0, 360.0)
            bucket['TWA'] = twa
        
        bucketStart = bucketEnd
        for key, value in bucket.items():
            # Make sure there's at least one valid data item for this bucket
            if key != 'ts' and value != None:
                l['samples'].append(bucket) # add this data item to the list of samples
                continue

    print("## analyze_race %s Leg %d samples %d" % (r['race'], leg+1, len(l['samples'])))

def average_sample_fields(samples, markBearing):
    avg_fields = ['AWA', 'AWS', 'STW', 'RUD', 'SOG', 'TWS', 'COG', 'HDG', 'TWD', 'TWA']
    d = {}
    counts = {}
    for field in avg_fields:
        d[field] = None
        counts[field] = 0

    for s in samples:
        for field in ['AWA', 'AWS', 'STW', 'RUD', 'SOG', 'TWS']:
            if s[field] != None:
                d[field] = s[field] if d[field] == None else d[field] + s[field]
                counts[field] += 1
        for field in ['COG', 'HDG', 'TWD', 'TWA']:
            if s[field] != None:
                # If course is biased north change range to [-180, 180]
                angle = s[field] if (markBearing >= 90 and markBearing <= 270) or (s[field] <= 180) else s[field] - 360
                d[field] = angle if d[field] == None else d[field] + angle
                counts[field] += 1
    for field in ['AWA', 'AWS', 'STW', 'RUD', 'SOG', 'TWS']:
        if d[field] != None:
            d[field] /= len(samples)
    for field in ['COG', 'HDG', 'TWD', 'TWA']:
        if d[field] != None:
            d[field] /= len(samples)
            d[field] += 0.0 if d[field] > 0 else 360.0 # Convert back to 0-360
    # Don't return a data bucket unless it has at least one valid data point
    for field in avg_fields:
        if d[field] != None:
            return(d)
    print("## average_sample_fields no data")
    return(None)

@dataclass
class LegItem:
    t: str = None
    comment: str = None
    start: datetime.datetime = None
    end: datetime.datetime = None
    duration: datetime.timedelta = None
    board: str = None
    hdg: float = None
    awa: float = None
    aws: float = None
    stw: float = None
    cog: float = None
    sog: float = None
    rud: float = None
    tws: float = None
    twd: float = None
    twa: float = None

# Return a list of analysis items. Each item can be a comment or a line of data.
# Data lines can be a minute, a board summary, or a leg summary
def analyze_by_minute(regatta, r):
    items = []
    c = regatta['courses'][r['course']]
    
    items.append(LegItem(t="Regatta", comment="%s - %s - %d race%s" % (regatta['boat'], regatta['regatta'], len(regatta['races']), "" if len(regatta['races']) < 2 else "s")))
    for leg, l in enumerate(r['legs']):
        markBearing = c["legs"][leg]["bearing"]
        items.append(LegItem(t="Leg", comment="Race %s Course %s Leg %d %s (%.2fnm @ %03d%s)" %
                             (r['race'], r['course'], leg+1, c["legs"][leg]["label"], c["legs"][leg]["distance"], markBearing, deg)))

    items.append(LegItem(t="Blank"))
    for leg, l in enumerate(r['legs']):
        legStart = l['startts']
        legEnd = l['endts']
        markBearing = c["legs"][leg]["bearing"]

        items.append(LegItem(t="Leg", comment="Race %s Leg %d %s (%.2fnm @ %03d%s) - (%s - %s) %s" %
                             (r['race'], leg+1, c["legs"][leg]["label"], c["legs"][leg]["distance"], markBearing, deg, addtz(legStart)[11:], addtz(legEnd)[11:], l['duration'])))
        
        legSamples = l['samples']
        if legSamples[0]['AWA'] == None:
            print("## No AWA at start of leg %s %r" % (addtz(legStart)[11:0], legSamples[0]))
        board = 'Port' if legSamples[0]['AWA'] < 0 else 'Stbd'
        # Loop through the AWAs looking for tacks and gybes
        l['boards'] = []
        bstart = l['startts']
        bend =  l['startts']
        for i in range(1, len(legSamples)):
            awa = legSamples[i]['AWA']
            # Use sample if not near a tack or gybe
            if abs(awa) > minAWA and abs(awa) < maxAWA:
                samplets = legSamples[i]['ts']
                if (board == 'Stbd' and awa < 0) or (board == 'Port' and awa > 0):
                    l['boards'].append((bstart, bend, board))
                    board = 'Port' if awa < 0 else 'Stbd'
                    #print("New board %s (%3f) @ %s" % (board, awa, samplets))
                    bstart = samplets
                bend = samplets # extend end timestamp to current sample - but not if tacking or gybing

        # Last board for this leg
        l['boards'].append((bstart, bend, board))
    
        print("## analyze_by_minute Leg %d %s #boards %d" % (leg, l['boards'][0][1], len(l['boards'])))
        
        bucketDelta = datetime.timedelta(seconds=reportSeconds)
        for b in range(len(l['boards'])):
            boardStart = l['boards'][b][0]
            boardEnd = l['boards'][b][1]
            items.append(LegItem(t="Board", comment="Race %s Leg %d Board %d @ %s - %s %s" % (r['race'], leg+1, b+1, addtz(boardStart)[11:], addtz(boardEnd)[11:], boardEnd - boardStart)))

            # Time range is start to end (not > start) since boardStart includes waiting for tack or gybe to finish
            boardSamples = [ s for s in legSamples if s['ts'] >= boardStart and s['ts'] <= boardEnd ]
            minute = {}
            bucketStart = boardStart
            while bucketStart < boardEnd:
                # Per-minute
                bucketEnd = bucketStart + bucketDelta
                bucketEnd = min(bucketEnd, boardEnd)
                bucketSamples = [ s for s in boardSamples if s['ts'] > bucketStart and s['ts'] <= bucketEnd ]
                savg = average_sample_fields(bucketSamples, markBearing)
                if savg != None:
                    items.append(LegItem("Minute", None, bucketStart, bucketEnd, None, l['boards'][b][2], savg['HDG'], savg['AWA'], savg['AWS'], savg['STW'], savg['COG'], savg['SOG'], savg['RUD'], savg['TWS'], savg['TWD'], savg['TWA']))
                bucketStart = bucketEnd
                    
            # Per-board
            if len(boardSamples) == 0:
                items.append(LegItem(t="Board", comment=("Board %s - %s %s - no samples" %
                                      (addtz(boardStart)[11:], addtz(boardEnd)[11:], boardEnd - boardStart))))
            else:
                savg = average_sample_fields(boardSamples, markBearing)
                if savg != None:
                    items.append(LegItem("Board", None, boardStart, boardEnd, boardEnd - boardStart, l['boards'][b][2], savg['HDG'], savg['AWA'], savg['AWS'], savg['STW'], savg['COG'], savg['SOG'], savg['RUD'], savg['TWS'], savg['TWD'], savg['TWA']))

                if (b < len(l['boards'])-1):
                    items.append(LegItem(t="Blank"))

        # Per-leg
        if len(legSamples) == 0:
            items.append(LegItem(t="Leg", comment=("Leg   %s - %s %s - no samples" %
                                                   (addtz(boardStart)[11:], addtz(boardEnd)[11:], boardEnd - boardStart))))
        else:
            savg = average_sample_fields(legSamples, markBearing)
            items.append(LegItem("Leg", None, legStart, legEnd, legEnd - legStart, None, savg['HDG'], savg['AWA'], savg['AWS'], savg['STW'], savg['COG'], savg['SOG'], savg['RUD'], savg['TWS'], savg['TWD'], savg['TWA']))
        items.append(LegItem("Blank"))
    return(items)

@dataclass
class XlsxFormats:
    degree3 = None
    float0 = None
    float1 = None
    float2 = None
    time = None
    timedelta = None
    bold = None
    board = None
    leg = None

xlsxF = XlsxFormats()

def per_race_xlsx(xl, regatta, r):
    items = analyze_by_minute(regatta, r)

    ws = xl.add_worksheet("%s_%s" % (regatta['regatta'], r['race']))

    column_formats = [ None, xlsxF.time, xlsxF.time, xlsxF.timedelta, None, xlsxF.degree3, xlsxF.degree3, xlsxF.float1, xlsxF.float1, xlsxF.degree3, xlsxF.float1, xlsxF.float1, xlsxF.float1, xlsxF.degree3, xlsxF.degree3 ]
    last_col = len(column_formats) - 1
    for col, f in enumerate(column_formats):
        if f != None:
            ws.set_column(col, col, None, f)
    
    row = 0
    column_labels = [ None, "Start", "End", "Duration", "Board", "HDG", "AWA", "AWS", "STW", "COG", "SOG", "RUD", "TWS", "TWD", "TWA" ]
    ws.set_row(row, None, xlsxF.header)
    ws.write_row(row, 0, column_labels)
    ws.freeze_panes(1, 0)
    row += 1
    
    for i in items:
        if i.t == "Leg" or i.t == "Board" or i.t == "Minute" or i.t == "Regatta":
            if i.comment != None:
                ws.write_row(row, 0, [i.t, i.comment])
            else:
                ws.write_row(row, 0, [i.t, addtz(i.start)[11:], addtz(i.end)[11:], i.end - i.start if (i.t == "Board") or (i.t == "Leg") else None, i.board if i.t == "Minute" else None,
                                      i.hdg, i.awa, i.aws, i.stw, i.cog, i.sog, i.rud, i.tws, i.twd, i.twa])
                if i.hdg == None and i.t == "Minute":
                    print("## %s HDG None" % (addtz(i.start)[11:]))
            row += 1
        elif i.t == "Blank":
            row += 1

    # Create the conditional highlighting for the regatta, leg, and board summary lines
    # I don't understand why we need the INDIRECT() call
    ws.conditional_format(0, 0, row, last_col, {'type': 'formula', 'criteria': '=INDIRECT("A"&ROW())="Regatta"', 'format': xlsxF.regatta})
    ws.conditional_format(0, 0, row, last_col, {'type': 'formula', 'criteria': '=INDIRECT("A"&ROW())="Leg"', 'format': xlsxF.leg})
    ws.conditional_format(0, 0, row, last_col, {'type': 'formula', 'criteria': '=INDIRECT("A"&ROW())="Board"', 'format': xlsxF.board})

def spreadsheet_report():
    bn = regattas[regattalist[0]]['boat']
    xlfn = "%s_%s.xlsx" % (bn, "aggregate" if len(regattalist) > 1 else regattas[regattalist[0]]['basefn'])

    with xlsxwriter.Workbook(xlfn) as xl:
        xlsxF.degree3 = xl.add_format({'num_format': '000'})
        xlsxF.float0 = xl.add_format({'num_format': '##0'})
        xlsxF.float1 = xl.add_format({'num_format': '##0.0'})
        xlsxF.float2 = xl.add_format({'num_format': '##0.00'})
        xlsxF.time = xl.add_format({'num_format': '[h]:mm:ss'})
        xlsxF.timedelta = xl.add_format({'num_format': '[h]:mm:ss'})
        xlsxF.header = xl.add_format({'bold': True, 'align': 'right'})

        # Pastels from https://www.schemecolor.com/rainbow-pastels-color-scheme.php
        xlsxF.regatta = xl.add_format({'bold': True, 'bg_color': '#FFB7B2'}) # Red-ish
        xlsxF.leg = xl.add_format({'bold': True, 'bg_color': '#E2F0CB'}) # Green-ish
        xlsxF.board = xl.add_format({'bg_color': '#C7CEEA'}) # Blue-ish

        for rn in regattalist:
            reg = regattas[rn]
            for r in reg['races']:
                per_race_xlsx(xl, reg, r)

def none_sub(v, f):
    # this is a kludgey way of returning a string the length of the format with a - at the end
    tmp = "     -"
    if v != None:
        return(f % (v))
    return(tmp[-int(f[1]):])

def per_leg_report(regatta, r):
    ofn = "%s_%s_%s_legs.txt" % (regatta['boat'], regatta['basefn'], r['race'])
    print("## Create %s" % (ofn))

    items = analyze_by_minute(regatta, r)
    
    # Create a minute-by-minute report for the leg
    with open(ofn, "w") as f:
        for i in items:
            if i.t == "Leg" or i.t == "Board" or i.t == "Minute" or i.t == "Regatta":
                if i.comment != None:
                    f.write("%s\n" % (i.comment))
                elif i.t == "Leg" or i.t == "Board":
                    #f.write("%6s %s - %s %8s HDG %3.0f AWA %4.0f AWS %4.1f STW %4.1f COG %3.0f SOG %4.1f RUD %5.1f TWS %4.1f TWD %4.0f TWA %4.0f\n" %
                    f.write("%6s %s - %s %8s HDG %s AWA %s AWS %s STW %s COG %s SOG %s RUD %s TWS %s TWD %s TWA %s\n" %
                            (i.t, addtz(i.start)[11:], addtz(i.end)[11:], i.end - i.start,
                             none_sub(i.hdg, "%3.0f"), none_sub(i.awa, "%4.0f"), none_sub(i.aws, "%4.1f"), none_sub(i.stw, "%4.1f"), none_sub(i.cog, "%3.0f"), none_sub(i.sog, "%4.1f"), none_sub(i.rud, "%5.1f"), none_sub(i.tws, "%4.1f"), none_sub(i.twd, "%4.0f"), none_sub(i.twa, "%4.0f")))
                elif i.t == "Minute":
                    #f.write("Minute %s - %s     %s HDG %3.0f AWA %4.0f AWS %4.1f STW %4.1f COG %3.0f SOG %4.1f RUD %5.1f TWS %4.1f TWD %4.0f TWA %4.0f\n" %
                    f.write("Minute %s - %s     %s HDG %s AWA %s AWS %s STW %s COG %s SOG %s RUD %s TWS %s TWD %s TWA %s\n" %
                            (addtz(i.start)[11:], addtz(i.end)[11:], i.board,
                             none_sub(i.hdg, "%3.0f"),
                             none_sub(i.awa, "%4.0f"),
                             none_sub(i.aws, "%4.1f"),
                             none_sub(i.stw, "%4.1f"),
                             none_sub(i.cog, "%3.0f"),
                             none_sub(i.sog, "%4.1f"),
                             none_sub(i.rud, "%5.1f"),
                             none_sub(i.tws, "%4.1f"),
                             none_sub(i.twd, "%4.0f"),
                             none_sub(i.twa, "%4.0f")))
                elif i.t == "Blank":
                    f.write("\n")

# Many possible line styles
# line_styles = cycle(['-','-','-', '--', '-.', ':', '.', ',', 'o', 'v', '^', '<', '>', '1', '2', '3', '4', 's', 'p', '*', 'h', 'H', '+', 'x', 'D', 'd', '|', '_'])

maxLegDuration = 0
maxPlotHeight = 3.0
maxPlotWidth = 10.0

def leg_chart(regatta, r, leg, fig, ax):
    y_scales = {}
    y_scales["compass"] = {"in_use": False, "min": 0, "max": 360, "label": ""}
    y_scales["apparent"] = {"in_use": False, "min": -180, "max": 180, "label": ""}
    y_scales["windspeed"] = {"in_use": False, "min": 0, "max": 25, "label": ""}
    y_scales["boatspeed"] = {"in_use": False, "min": 0, "max": 12, "label": ""}
    y_scales["rudder"] = {"in_use": False, "min": -15, "max": 15, "label": ""}

    plotItems = {}
    plotItems['COG'] = { 'label': 'COG', 'scale': 'compass', 'color': 'black', 'style': '-' }
    plotItems['SOG'] = { 'label': 'SOG', 'scale': 'boatspeed', 'color': 'black', 'style': '-' }
    plotItems['TWD'] = { 'label': 'TWD', 'scale': 'compass', 'color': cmap[0], 'style': '--' }
    plotItems['TWS'] = { 'label': 'TWS', 'scale': 'windspeed', 'color': cmap[0], 'style': '-' }
    plotItems['AWA'] = { 'label': 'AWA', 'scale': 'apparent', 'color': cmap[1], 'style': '-' }
    plotItems['AWS'] = { 'label': 'AWS', 'scale': 'windspeed', 'color': cmap[1], 'style': '--' }
    plotItems['HDG'] = { 'label': 'HDG', 'scale': 'compass', 'color': cmap[4], 'style': '-' }
    plotItems['STW'] = { 'label': 'STW', 'scale': 'boatspeed', 'color': cmap[2], 'style': '-' }
    plotItems['RUD'] = { 'label': 'Rudder', 'scale': 'rudder', 'color': cmap[3], 'style': ':' }

    c = regatta['courses'][r['course']]
    l = r['legs'][leg]

    legDataFields = ["AWA", 'STW', "TWD", "TWS", "RUD"]
    legData = []
    for ld in legDataFields:
        if l['samples'][0][ld] != None:
            legData.append(ld)

    firstTime = l['startts'] + tzoffset
    lastTime = firstTime + datetime.timedelta(seconds=maxLegDuration)

    print("## leg_chart Race %s Leg %r" % (r['race'], leg))
    
    ax[leg].set_title(loc='left', label="Race %s Leg %d %s (%.2fnm @ %03d%s) - (%s - %s) %s"
                      % (r['race'], leg+1, c["legs"][leg]["label"], c["legs"][leg]["distance"], c["legs"][leg]["bearing"], deg, l['start'][11:], l['end'][11:], l['duration']))
    ax[leg].set_xlim(firstTime, lastTime)
    ax[leg].hlines(y=0, color='black', xmin=firstTime, xmax=lastTime, linestyles='--', linewidth=0.5)

    hostUsed = False
    for d in legData:
        s = y_scales[plotItems[d]["scale"]]
        if not s["in_use"]:
            s["in_use"] = True
            s["label"] = plotItems[d]["label"]
            if not hostUsed:
                s["isHost"] = True
                s["axis"] = ax[leg]
            else:
                s["isHost"] = False
                s['axis'] = ax[leg].twinx()
            hostUsed = True
        else:
            s["label"] = "%s, %s" % (s["label"], plotItems[d]["label"])

    numScales = 0
    for scale in y_scales:
        if s["in_use"]:
            numScales += 1
    scaleWidth = 1.2 # inches for each scale

    scalesWidth = numScales * scaleWidth
    maxDataWidth = maxPlotWidth - scalesWidth
    scalesPct = (scalesWidth-1) /  maxPlotWidth # one scale is on the left

    i = 0
    for scale in y_scales:
        s = y_scales[scale] 
        if s["in_use"]:
            #print("Axis R%s L%d Axis %d: %s(%s)" % (race, leg, i, scale, s['label']))
            a = s["axis"]
            a.set_ylabel(s["label"])
            a.set_autoscaley_on(False)
            a.set_ylim(s["min"], s["max"])
            if not s['isHost']:
                a.spines['right'].set_position(('axes', 1.0+(scalesPct * i / numScales)))
                i += 1

    legendLines = []
    portLegend = None
    stbdLegend = None
    for d in legData:
        if (l['samples'][0][d] == None):
            # Don't plot if there's no data
            continue
        a = y_scales[plotItems[d]['scale']]['axis']

        x = [ (s['ts']+tzoffset) for s in l['samples'] ]
        y = [ s[d] for s in l['samples'] ]
        color = plotItems[d]['color']
        a.yaxis.label.set_color(color)
        a.tick_params(axis='y', colors=color)
        a.spines['right'].set_color(color)
        #if d == 'TWS(avg)':
        #    a.hlines(y=[12, 15, 18], color=color, xmin=firstTime, xmax=lastTime, linestyles=':', linewidth=0.5)
                    
        style = plotItems[d]['style']
        #print("Plot R%s L%d B%d %s %s %s" % (race, leg, b, d, color, style))
        smooth = scipy.signal.savgol_filter(y, 7, 3)
        line, = a.plot(x, smooth, color=color, linestyle=style, label=plotItems[d]['label'])
        legendLines.append(line)

    # Setting the xaxis has to come after the plot - maybe because it needs x data?
    xfmt = mdates.DateFormatter("%H:%M")
    ax[leg].xaxis.set_major_formatter(xfmt)

    if portLegend:
        legendLines.insert(0, portLegend)
    if stbdLegend:
        legendLines.insert(0, stbdLegend)
    ax[leg].legend(handles=legendLines, loc='best')

    #plt.subplot(ax[leg])
    #plt.axhline()
    #plt.close(fig)

def strip_charts(regatta, r):
    legs = len(r['legs'])
    #fig, ax = plt.subplots(legs, figsize=(maxPlotWidth, maxPlotHeight * legs), constrained_layout=True)
    #fig, ax = plt.subplots(legs, constrained_layout=True)
    #fig, ax = plt.subplots(legs)

    fig, ax = plt.subplots(legs, figsize=(maxPlotWidth, maxPlotHeight * legs + .5))
    #fig.autofmt_xdate()
    fig.suptitle("%s - %s Race %s - %s %s - %s\n " % (regatta['boat'], regatta['regatta'], r['race'], r['legs'][0]['start'][:10], r['legs'][0]['start'][11:], r['legs'][-1]['end'][11:]))

    # Find the longest leg to make all the strip charts the same length
    maxDur = datetime.timedelta(seconds=15 * 60)
    for l in r['legs']:
        dur = l['endts'] - l['startts']
        maxDur = dur if maxDur < dur else maxDur
    # Round to the next minute for short legs or 5 minutes for longer races
    minutes = int(maxDur.total_seconds() / 60)
    minutes += 1 if minutes < 30 else 5
    global maxLegDuration
    maxLegDuration = minutes * 60

    for leg in range(len(r['legs'])):
        leg_chart(regatta, r, leg, fig, ax)

    #fig.set_constrained_layout_pads(hspace=5.0, h_pad=1.0)
    plt.tight_layout(h_pad=0.0, w_pad=0.0)
    plt.subplots_adjust(top=0.94, bottom=0.04, left=0.08, right=0.72, wspace=0.05, hspace=0.25)
    #plt.subplots_adjust(top=0.96, bottom=0.04, left=0.08, right=0.72, wspace=0.05, hspace=0.25)
    #plt.show()
    plt.savefig("%s_%s_%s_strip" % (regatta['boat'], regatta['basefn'], r['race']), bbox_inches="tight")

################### Polar variables

# polarData defines the wind ranges we're interested in graphing. Should probably be defined externally.

# This range matched an ORRez certificate
polarData = [
    { 'min':  0.0, 'max':  9.3, 'ax': None, 'data': [] },
    { 'min':  9.3, 'max': 11.5, 'ax': None, 'data': [] },
    { 'min': 11.5, 'max': 14.6, 'ax': None, 'data': [] },
    { 'min': 14.6, 'max': 18.7, 'ax': None, 'data': [] },
    { 'min': 18.7, 'max': 22.0, 'ax': None, 'data': [] },
    { 'min': 22.0, 'max': 28.0, 'ax': None, 'data': [] }
]

# This range matches where a particular boat accelerates
# Six plots fit well on a page
polarData = [
    { 'min':  0, 'max': 11, 'ax': None, 'data': [] },
    { 'min': 11, 'max': 13, 'ax': None, 'data': [] },
    { 'min': 13, 'max': 15, 'ax': None, 'data': [] },
    { 'min': 15, 'max': 18, 'ax': None, 'data': [] },
    { 'min': 18, 'max': 22, 'ax': None, 'data': [] },
    { 'min': 22, 'max': 28, 'ax': None, 'data': [] }
]

def gather_polar_data(rl):
    for reg in rl:
        regatta = regattas[reg]
        print("## plot_polar regatta %s" % (regatta['regatta']))
        for r in regatta['races']:
            print("## plot_polar regatta %s race %s" % (regatta['regatta'], r['race']))
            racecount += 1
            legcount += len(r['legs'])

            for leg, l in enumerate(r['legs']):
                for s in l['samples']:
                    # Add this sample to the correct polar wind range
                    twa = s['TWA']
                    tws = s['TWS']
                    stw = s['STW']
                    twd = s['TWD']
                    hdg = s['HDG']
                    ts = s['ts']
                    
                    for p, polar in enumerate(polarData):
                        # If the max is greater than tws then this is the right bucket
                        if polar['max'] > tws:
                            theta = radians(twa)
                            datum = (theta, stw, 'red' if theta > pi else 'green', r['race'], leg, ts, twd, hdg, twa)
                            #datum = (theta, stw, plt.cm.Set2.colors[p], int(race), leg, l['Time'][d], l['TWD(med)'][d], l['Heading'][d])
                            polar['data'].append(datum)
                            break
        
lineSpan = 7.5

# Plot performance for various wind ranges and
# Make an aggregate plot with all wind ranges
# Six plots fit well on a standard page
def plot_polars():
    fig, axes = plt.subplots(2, 3, figsize=(10, 8), subplot_kw=dict(polar=True), constrained_layout=True)
    polarData[0]['ax'] = axes[0, 0]
    polarData[1]['ax'] = axes[0, 1]
    polarData[2]['ax'] = axes[0, 2]
    polarData[3]['ax'] = axes[1, 0]
    polarData[4]['ax'] = axes[1, 1]
    polarData[5]['ax'] = axes[1, 2]

    regattacount = len(regattas)
    racecount = 0
    legcount = 0
    bn = regattas[regattalist[0]]['boat']
    if regattacount > 1:
        fig.suptitle("%s - %d Regatta%s, %d Race%s, %d Legs\n" % (bn,
                                                                  regattacount,
                                                                  "" if regattacount == 1 else 's',
                                                                  racecount,
                                                                  "" if racecount == 1 else 's',
                                                                  legcount))
    else:
        fig.suptitle("%s - %s, %d Race%s, %d Legs\n" % (bn,
                                                        regattas[regattalist[0]]['regatta'],
                                                        racecount,
                                                        "" if racecount == 1 else 's',
                                                        legcount))

    for p, pd in enumerate(polarData):
        pd['ax'].set_title("%4.1fkts < Wind < %4.1fkts\n%d samples\n " % (pd['min'], pd['max'], len(pd['data'])), va='bottom')
        print("## %f < Wind < %f - %d samples" % (pd['min'], pd['max'], len(pd['data'])))
        ax = pd['ax']
        ax.set_thetalim(thetamin=-180, thetamax=180)
        ax.set_ylim(0, 12)
        ax.set_theta_offset(pi/2.0)
        ax.set_theta_direction(-1)
        theta_ticks = [(float(t)/360.0) * tau for t in [-135, -90, -45, 0, 45, 90, 135, 180]]
        ax.set_xticks(theta_ticks)

        if len(pd['data']) > 0:
            t, s, c, r, leg, time, twd, hdg, twa = zip(*pd['data']) # unzip the data to theta, speed, color w/ magic * operator
            ax.scatter(t, s, color=c, marker='.', alpha=0.05)

        # Sort by theta
        pd['data'].sort(key=lambda x: x[0])
        points = []

        # Compute a simple per-bucket mean, but keep a list of STWs so we can use stats on it
        bucketSTW = []
        bucketTheta = 0
        bucketCount = 0
        bucket = minTWA
        next_bucket = bucket + lineSpan
        for d in pd['data']:
            (theta, stw, color, raceName, leg, timeStamp, twd, hdg, twa) = d
            while twa > next_bucket:
                if len(bucketSTW) != 0:
                    btheta = bucketTheta/bucketCount
                    bmean = mean(bucketSTW)
                    p90 = np.percentile(np.array(bucketSTW), 90)
                    #print("## Bucket [%d] %6.1f%s (Next %4.2f%s): %d" % (len(points), bucket, deg, next_bucket, deg, bucketCount))
                    points.append((btheta, bmean, p90, degrees(btheta)))
                bucket = next_bucket
                next_bucket += lineSpan
                bucketSTW = []
                bucketTheta = 0
                bucketCount = 0
            bucketSTW.append(stw)
            bucketTheta += theta
            bucketCount += 1

        # Last bucket
        if len(bucketSTW) != 0:
            btheta = bucketTheta/bucketCount
            bmean = mean(bucketSTW)
            p90 = np.percentile(np.array(bucketSTW), 90)
            #print("## Bucket [%d] %6.1f%s (Last): %4.1f %3.0f" % (len(points), bucket, deg, theta, degrees(theta)))
            points.append((btheta, bmean, p90, degrees(btheta)))

        # make the polar lines rounder by interpolating missing points
        i = 0
        while i < len(points)-1:
            (t0, s0, p0, b0) = points[i]
            (t1, s1, p1, b1) = points[i+1]
            bprime = b0 + lineSpan
            if (bprime < b1) and ((b0 >= minTWA and b0 <= maxTWA) or (b0 >= 360-maxTWA and b0 <= 360-minTWA)) and ((b1 >= minTWA and b1 <= maxTWA) or (b1 >= 360-maxTWA and b1 <= 360-minTWA)):
                # Slope of line * degrees of run
                bmeanprime = s0 + (((s1 - s0) / (b1 - b0)) * (bprime - b0))
                p90prime = p0 + (((p1 - p0) / (b1 - b0)) * (bprime - b0))
                points.insert(i+1, (radians(bprime), bmeanprime, p90prime, bprime))
                #print("## Insert [%d] %6.2f < %6.2f < %6.2f (%4.1f & %4.1f)" % (i, b0, bprime, b1, bmeanprime, p90prime))
            i += 1

        # Look for errors in data order (this is looking for bugs) - the above code was fussy
        for i in range(len(points)-1):
            theta0 = points[i][0]
            bucket0 = points[i][3]
            theta1 = points[i+1][0]
            bucket1 = points[i+1][3]
            if  theta0 >= theta1:
                print("## Backwards [%d]: %4.2f -> %4.2f (%5.2f -> %5.2f) (%3.0f -> %3.0f)" % (i, theta0, theta1, bucket0, bucket1, degrees(theta0), degrees(theta1)))

        if len(points) > 1:
            # First work the starboard side
            for scloseHauled in range(0, len(points)):
                twa = degrees(points[scloseHauled][0])
                #print("## Close Hauled Stbd[%d] %f - %4.1f vs %4.1f" % (scloseHauled, points[scloseHauled][0], twa, minTWA))
                if twa >= minTWA:
                    break
            
            for sgybing in range(scloseHauled, len(points)):
                twa = degrees(points[sgybing][0])
                #print("## Gybing       Stbd[%d] %f - %4.1f vs %4.1f" % (sgybing, points[sgybing][0], twa, maxTWA))
                if twa > maxTWA:
                    break

            # Then work the port side
            for pgybing in range(sgybing, len(points)):
                twa = degrees(points[pgybing][0]) - 360
                #print("## Gybing       Port[%d] %f - %4.1f vs %4.1f" % (pgybing, points[pgybing][0], twa, -maxTWA))
                if twa > -maxTWA:
                    break

            for pcloseHauled in range(pgybing, len(points)):
                twa = degrees(points[pcloseHauled][0]) - 360
                #print("## Close Hauled Port[%d] %f - %4.1f vs %4.1f" % (pcloseHauled, points[pcloseHauled][0], twa, -minTWA))
                if twa > -minTWA:
                    break
            
            if scloseHauled != sgybing:
                stheta, speed, p90, bucket = zip(*(points[scloseHauled:sgybing]))
                smooth = scipy.signal.savgol_filter(speed, 7, 3)
                ax.plot(stheta, smooth, color='orange', linestyle='-')
                ssmooth = scipy.signal.savgol_filter(p90, 7, 3)
                ax.plot(stheta, ssmooth, color='blue', linestyle='-')

            if pcloseHauled != pgybing:
                ptheta, speed, p90, bucket = zip(*(points[pgybing:pcloseHauled]))
                smooth = scipy.signal.savgol_filter(speed, 7, 3)
                ax.plot(ptheta, smooth, color='orange', linestyle='-')
                psmooth = scipy.signal.savgol_filter(p90, 7, 3)
                ax.plot(ptheta, psmooth, color='blue', linestyle='-')

            # Save these lines for drawing combined polar
            pd['p90'] = ((ptheta, psmooth, stheta, ssmooth))

    pn = "%s_%s_polars" % (bn, "aggregate" if len(regattalist) > 1 else regattas[regattalist[0]]['basefn'])
    plt.savefig(pn, bbox_inches="tight")

    # One plot with all wind ranges
    fig, ax = plt.subplots(1, 1, figsize=(8, 10), subplot_kw=dict(polar=True), constrained_layout=True)
    if regattacount > 1:
        fig.suptitle("%s - %d Regatta%s, %d Race%s, %d Legs\n" % (bn,
                                                                  regattacount,
                                                                  "" if regattacount == 1 else 's',
                                                                  racecount,
                                                                  "" if racecount == 1 else 's',
                                                                  legcount))
    else:
        fig.suptitle("%s - %s, %d Race%s, %d Legs\n" % (bn,
                                                        regattas[regattalist[0]]['regatta'],
                                                        racecount,
                                                        "" if racecount == 1 else 's',
                                                        legcount))
    ax.set_thetalim(thetamin=-180, thetamax=180)
    ax.set_ylim(0, 12)
    ax.set_theta_offset(pi/2.0)
    ax.set_theta_direction(-1)
    theta_ticks = [(float(t)/360.0) * tau for t in [-135, -90, -45, 0, 45, 90, 135, 180]]
    ax.set_xticks(theta_ticks)

    for p, pd in enumerate(polarData):
        (ptheta, psmooth, stheta, ssmooth) = pd['p90']
        ax.plot(ptheta, psmooth, color=cmap[p], linewidth=2, linestyle='-', label="%2.0f kts" % (pd['min'] + ((pd['max'] - pd['min']) / 2)))
        ax.plot(stheta, ssmooth, color=cmap[p], linewidth=2, linestyle='-')
    plt.legend(loc='best')
    pn = "%s_%s_combined_polars" % (bn, "aggregate" if len(regattalist) > 1 else regattas[regattalist[0]]['basefn'])
    plt.savefig(pn, bbox_inches="tight")

# Generate text polars file suitable for Expedition
def expedition_polars():
    bn = regattas[regattalist[0]]['boat']
    en = "%s_%s_polars.txt" % (bn, "aggregate" if len(regattalist) > 1 else regattas[regattalist[0]]['basefn'])
    with open(en, "w") as f:
        f.write("!Expedition polar - %s\n" % (bn))
        for p in range(len(polarData)):
            pd = polarData[p]
            f.write("%-4.1f" % ((pd['min'] + pd['max']) / 2.0))
            points = []
            for center in range(45, 181, 15):
                c = float(center)
                high = c - 7.5
                low = c + 7.5
                for d in pd['data']:
                    (theta, stw, color, r, leg, time, twd, hdg, twa) = d
                    twa = abs(twa)
                    if twa > high and twa <= low:
                        points.append(stw)
                f.write("  %4.1f %5.2f" % (float(center), np.percentile(np.array(points), 90) if len(points) > 0 else 0))
            f.write("\n")

# Create a gpx track file. Could add waypoints for marks, tacks & gybes, etc. Could annotate w/ sensor data
def gpx_track(regatta, r):
    ofn = "%s_%s_%s.gpx" % (regatta['boat'], regatta['basefn'], r['race'])
    with open(ofn, "w") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('<gpx xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" version="1.1"\n')
        f.write('creator="polarize"\n')
        f.write('xmlns="http://www.topografix.com/GPX/1/1"\n')
        f.write('xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd">\n')
        now = datetime.datetime.now(datetime.timezone.utc)
        f.write('<metadata><time>%s</time></metadata>\n' % (now.strftime("%Y-%m-%dT%H:%M:%SZ"))) # 2019-09-13T21:17:14Z
        f.write('<trk>\n')
        f.write('  <name>%s_%s_%s</name>\n' % (regatta['boat'], regatta['basefn'], r['race']))
        f.write('  <trkseg>\n')

        last_ts = datetime.datetime(year=1970, month=1, day=1, hour=0, minute=0, second=0)
        for p in r['LATLON']:
            if p[0] != last_ts:
                tstring = p[0].strftime("%Y-%m-%dT%H:%M:%SZ")
                f.write('    <trkpt lat="%.6f" lon="%.6f"><time>%s</time></trkpt>\n' % (p[1], p[2], tstring))
                last_ts = p[0]

        f.write('  </trkseg>\n')
        f.write('</trk>\n')
        f.write('</gpx>\n')
#                f.write('<wpt lat="%.5f" lon="%.5f"><name>%s</name></wpt>\n' % (p[1], p[2], tstring))

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate polars from boat data', epilog='Each regatta is usually in its own directory')
    parser.add_argument("-strip", default=False, action='store_true', dest='strip', help='Create per-race strip graph files')
    parser.add_argument("-legs", default=False, action='store_true', dest='leg', help='Create per-leg analysis')
    parser.add_argument("-minute", default=False, action='store_true', dest='leg', help='Create minute-by-minute reports in per-leg analysis')
    parser.add_argument("-spreadsheet", default=False, action='store_true', dest='spreadsheet', help='Create xlsx spreadsheet file')
    parser.add_argument("-polars", default=False, action='store_true', dest='polars', help='Create aggregate polar graph file')
    parser.add_argument("-exp", default=False, action='store_true', dest='exp', help='Create aggregate Expedition polar text file')
    parser.add_argument("-gpx", default=False, action='store_true', dest='gpx', help='Create gpx track')
    parser.add_argument('regatta', nargs='*', help='Regatta JSON description files')
    args = parser.parse_args()

    for arg in args.regatta:
        parse_regatta(arg)

    regattalist.sort()
    for rn in regattalist:
        reg = regattas[rn]
        tzoffset = datetime.timedelta(hours=reg['tz']) # global for this regatta
        
        for race in reg["races"]:
            parse_race(reg, race)
            analyze_race(reg, race)

            for leg in range(len(race["legs"])):
                analyze_leg(reg, race, leg)

            if args.leg:
                per_leg_report(reg, race)

            if args.strip:
                strip_charts(reg, race)

            if args.gpx:
                gpx_track(reg, race)

    if args.polars or args.exp:
        gather_polar_data(regattalist)

        if args.polars:
            plot_polars()

        if args.exp:
            expedition_polars()

    if args.spreadsheet:
        spreadsheet_report()
