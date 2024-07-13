#!/usr/bin/env python3


import pprint
import sys

pp = pprint.PrettyPrinter(indent=4)


def main():
	filename="input files/test_sample.txt"
	with open(filename) as file:
		while raw_line := file.readline():
			# We need to examine the line to see if there is a FORM FEED (0x0C) within it, if so
			# the line needs to be split on that character and each half treated as a separate line
			# The W11 print to Generic Text or the Quantum software inserts FFs at the end of the page
			raw_line=raw_line.rstrip()
			lines=raw_line.split('\x0c')
			for line in lines:
				process_line(line)

def	process_line(line):
	#print(line.rstrip())
	if len(line)>0 and line[0].isnumeric():		# Data sample lines are the only ones starting with a digit
		parse_line(line.rstrip())
		#sys.exit(0)

def	parse_line(line):

	time_position=0
	time_length=8
	date_position=10
	date_length=10
	remainder_position=20
	speed_position=28
	speed_length=4
	tmc_position=32
	tmc_length=4
	
	print("Sanity checking not done yet!")
	print(line)
	# Date and time are in fixed positions starting at column 0 and in the format
	# hh:mm:ss- mm/dd/yyyy
	time=line[time_position:time_position+time_length]
	print("time = " + time)
	date_us_fmt=line[date_position:date_position+date_length]
	print("date = " + date_us_fmt) 
	# Because the mileage field leading space is lost when the distance goes to 3 figures and
	# extends when it goes to 4 figures, we start at the end of the date field then strip any leading
	# spaces. The mileage field should be followed by a space as the next field will be the speed
	parts=line[remainder_position:].lstrip().split(" ",1)
	mileage=float(parts[0])
	print("mileage = " + str(mileage))
	# We need to handle speed and tmc carefully as the TMC field will lose leading spaces when it
	# goes over 3 digits
	#pp.pprint(line[28:32])
	speed=int(line[speed_position:speed_position+speed_length].strip())
	print("speed = "+str(speed))
	tmc=int(line[tmc_position:tmc_position+tmc_length].strip())
	print("tmc = "+str(tmc))
	parts=line[speed_position+speed_length+tmc_length:].lstrip().split()
	#pp.pprint(parts)
	
	# The first 3 fields are values as follows:
	brake_pipe_pressure=int(parts[0])
	print("BP = "+str(brake_pipe_pressure))
	brake_cylinder_pressure=int(parts[1])
	print("BC = "+str(brake_cylinder_pressure))
	throttle_position=parts[2]			# This is left as string to cater for (D)ynamic or low (ID)le states
	print("TP = "+throttle_position)
	
	# The rest of these are binary flags (1 is on, 0 is off)
	# Flags are: 
	# Reverse
	# EIE (Engineer induced emergency)
	# Pressure Control Switch (set when BP < 45 psi)
	# Headlight - short end
	# Forward
	# Headlight - long end
	# Horn
	# Digital Spare 1
	# Digital Spare 2
	# Vigilance Control Alert Acknowledged
	# Axle Drive TypeError
	flags=parts[3:]
	pp.pprint(flags)
	
	
	
	
if __name__ == '__main__':
    main()
