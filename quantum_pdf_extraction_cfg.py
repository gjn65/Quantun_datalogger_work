# File extacted from Quantum Desktop Software
# - print to PDF file to generate.
source_file='844_download.pdf'

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
