# File extacted from Quantum Desktop Software
# - print to PDF file to generate.
source_file='input files/844_20230712_CT3.pdf'
#source_file='sample1.pdf'
#source_file='sample2.pdf'

# Wheel diameter in mm - this may be used to correct the speed calculated by the
# QDP software which uses a figure embedded in the logger (which will be in inches)
wheel_dia_actual_mm = 995

# Speed adjustment factor.
# This will be the actual wheel diameter divided by the QDP reported diameter (converted to mm)
# If this sis set to 0 then the calculation has not yet been made and the code will assume a 1:1 ratio
# The QDP wheel size is generally reported on Page 0 of the printout with a pair of lines as follows:
# "Wheel size used by program" followed immediately by the line
# Circumference = xxx.x Diameter = xx.x
speed_adjustment_factor=0

# Any line containing one of these phrases in omitted from processing
# (as are all lines on Page 0 - with the exception of the wheel diameter line)
skip_list_words = ["Quantum Desktop Playback",
                   "Report Date",
                   "Locomotive",
                   'TIME']

# Workbook name - including path if required - no xlsx suffix, that is added by the code
workbook_name='output/qdp_output'
worksheet_name="Data Extract"

# Required date range.
# Define the start and end date/times as yyyy/mm/dd hh:mm:ss
# Only records between these timestamps will be reported.
# The between_dates flag is set to True to activate this test, or False to ignore it.
filter_dates=True
start_timestamp="2023/07/12 00:00:00"
end_timestamp="2023/07/12 23:59:59"

# This dictionary translates the throttle position value to a meaningful text.
# Note that Idle is stored in the logger output as "ID" but I modify it to "I"
# for parsing reasons
tp_translations={"F":"Fault","I":"Idle","ID":"Idle","D":"Dyn","S":"Stop"}

# Column headers for worksheet
# This is a list of tuples, each tuple has the column header text and a boolean flag indicating
# whether or not the column is visible
headers=[("Date",True),
         ("Time",True),
         ("Kilometres",True),
         ("Speed (kph)",True),
         ("TMC (A)",True),
         ("ABrk (psi)",True),
         ("IBrk (psi)",True),
         ("Throttle",True),
         ("Reverse",True),
         ("EIE",False),
         ("PCS",True),
         ("Light (S)",True),
         ("Forward",True),
         ("Light (L)",True),
         ("Horn",True),
         ("DS1",False),
         ("DS2",False),
         ("VS Ack",False),
         ("Axle Drive",False),
         ]

# Worksheet protection string
protect_string="3801"
protection_mode={'select_locked_cells':True,"select_unlocked_cells":True,"sort":True,"autofilter":True}


