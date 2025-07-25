# File extracted from Quantum Desktop Software
# - print to Generic Text file to generate.
source_file = 'input files/JULY2025.prn'
#source_file = 'input files/20231119 to 20240721.txt'
#source_file = 'input files/testinput.txt'
# source_file='sample1.pdf'
# source_file='sample2.pdf'

# The logger realtime clock can vary from the actual time so the timestamps
# are not accurate. A function will normalise the timestamps based on the
# TOD adjustment factor. If the logger clock is behind the real clock then a
# positive adjustment is made, if ahead then a negative adjustment is made.
# The adjustment factor is in seconds.
ts_adjustment = 0

# Wheel diameter in mm - this may be used to correct the speed calculated by the
# QDP software which uses a figure embedded in the logger (which will be in inches)
# The combination of this value and the wheel diameter reported in the data logger
# output is used to derive the speed adjustment factor.
#
# There will be multiple instances of this entry, one per locomotive known to the system.
# The values are stored in a dictionary, keyed with the locomotive number (as text)
# If there is no valid var for the named loco then the script will stop.
wheel_dia_actual_mm = {"844": 995, "845": 995}
#wheel_dia_actual_mm = 995 - old version prior to 2025/02/13 change

# Speed adjustment factor.
# This will be the actual wheel diameter divided by the QDP reported diameter (converted to mm)
# If this sis set to 0 then the calculation has not yet been made and the code will assume a 1:1 ratio
# The QDP wheel size is generally reported on Page 0 of the printout with a pair of lines as follows:
# "Wheel size used by program" followed immediately by the line
# Circumference = xxx.x Diameter = xx.x
# NB: This will be calculated in the code, any non-zero setting made manually is overwritten.
speed_adjustment_factor = 0

# Any line containing one of these phrases in omitted from processing
# (as are all lines on Page 0 - except for the wheel diameter line)
skip_list_words = ["Quantum Desktop Playback",
                   "Report Date",
                   "Locomotive",
                   'TIME']

# Workbook name - including path if required - no xlsx suffix, that is added by the code
# 				  as is the loco name and the date
workbook_name = 'output/qdp_output'
worksheet_name = "Data Extract"

# Required date range.
# Define the start and end date/times as yyyy/mm/dd hh:mm:ss
# Only records between these timestamps will be reported.
# The between_dates flag is set to True to activate this test, or False to ignore it.
filter_dates = True
start_timestamp = "2025/07/16 00:00:00"
end_timestamp = "2025/07/16 23:59:59"

# The data logger TOD clock is reverting to 1990 from time to time leading to
# oddball sample times in the traces. If this flag is set to True then these
# samples will be accepted for processing, if it is set to False then the samples
# will  be ignored.
epoch_timestamps_allowed = True
epoch_year = 1990

# This dictionary translates the throttle position value to a meaningful text.
# Note that Idle is stored in the logger output as "ID" but I modify it to "I"
# for parsing reasons
tp_translations = {"F": "Fault", "I": "Idle", "ID": "Idle", "D": "Dyn", "S": "Stop"}

# Column headers for worksheet
# This is a list of tuples, each tuple has the column header text and a boolean flag indicating
# whether the column is visible
headers = [("Date", True),
           ("Time", True),
           ("Kilometres", True),
           ("Speed (kph)", True),
           ("TMC (A)", True),
           ("ABrk (psi)", True),
           ("IBrk (psi)", True),
           ("Throttle", True),
           ("Reverse", True),
           ("EIE", False),
           ("PCS", True),
           ("Light (S)", True),
           ("Forward", True),
           ("Light (L)", True),
           ("Horn", True),
           ("DS1", False),
           ("DS2", False),
           ("VS Ack", False),
           ("Axle Drive", False),
           ]

# Used to sanity check the input and ensure we have the printing set up correctly in the Quantum Desktop Software
number_of_flags_expected=11

# Worksheet protection string
protect_string = "3801"
protection_mode = {'select_locked_cells': True, "select_unlocked_cells": True, "sort": True, "autofilter": True}

# In flight analysis code related variables
# Monitors incidents where the TMC exceeds a threshold - defined below - whilst the throttle is in the IDLE position.
# If such an event is detected then that event plus the events leading up to it are written to a page in the
# workbook. The number of lead-in events and the TMC threshold are defined below.
# The detection can be switched off by setting the enabled switch to False.
in_flight_analysis_enabled = True
ifa_deque_maxlen = 10               # Number of events to store in the queue
ifa_tmc_threshold = 0               # Extract records with TMC values exceeding this value
ifa_in_event_of_interest = False    # Set to true when we are in a run of records to be logged to the output file - flag is maintained by the code, not user set

# Hide stationary loco events
# If set to True, events with a speed of 0 kph and tmc - 0 amps and throttle position in idle are hidden - the 0 speed event leading into and exiting from
# such a run of stationary events are written to the spreadsheet with a notation written between them to indicate
# the case.
# If set to False then all events are written out
suppress_stationary_events=True