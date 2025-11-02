#!/Library/Frameworks/Python.framework/Versions/3.10/bin/python3.10


"""

Quantum Desktop Playback - data reporter

This code will parse a text file and create an Excel workbook.
The text file is created by running the QDP software, selecting the timescale
required using tags (or select the entire file) then printing it to a GENERIC/TEXT file
which must be in landscape format

NB: Ensure printer properties are set to LANDSCAPE mode for text output

When setting up the printed page, select ALL variables logged - this code will weed out those
that are not required.

The header page of the report is used to make adjustments to the wheel speed from that
recorded in the data (which can be modified if the data is extracted using Quantum Desktop Playback but cannot if
it's downloaded using the Quantum Recorder Service Toolkit code).

All logger data is in imperial units, so the code will convert to metric in the case of speed and distance. BP pressures
are left in imperial units as that is what is used in SteamRanger - this can be over-ridden if required with a
switch in the configuration file.

The cfg file contains all the data required to drive the application operation.

The workbook is locked by default, the password is im the cfg file.

-------------------------------------------------------------------------------------------------------------------------------

Command line arguments - as of 11/08/2025
-----------------------------------------
Switches                    Details                         Effect
-f --filename               source file path                overrides cfg.source_file
-a --ts_adjust              brings logger clock into sync   overrides cfg.ts_adjust
                            with real time if required.
                            Value is in seconds and is
                            relative to the logger clock
                            - a +ve value advances the
                            logger clock towards the rtc
-b --begin_timestamp        Filters records by date.        over-rides cfg.start_timestamp and forces cfg.filter_dates to true
                            Must be of the form
                            "yyyy/mm/dd hh:mm:ss"
                            End timestamp must also be
                            supplied
-e --end_timestamp          Filters records by date.        over-rides cfg.end_timestamp and forces cfg.filter_dates to true
                            Must be of the form
                            "yyyy/mm/dd hh:mm:ss"
                            Start timestamp must also be
                            supplied
-k --kpa_pressure           Report pressures in kpa         over-rides cfg.report_kpa_pressures
-p --psi_pressures          Report pressures in psi         over-rides cfg.report_kpa_pressures
-s --suppress_stationary    Suppress events where loco      over-rides cfg.suppress_stationary_events
                            is stationary with 0 TMC and
                            throttle in idle
-n --no_suppress_stationary Do not suppress events          over-rides cfg.suppress_stationary_events
                            where loco is stationary with
                            TMC = 0 amps and
                            throttle in idle
-i --integer_idle           If set, IDLE throttle position records will be reported as integer 0 in the spreadsheet
                            to facilitate adding charts. Over-rides the config file entry idle_as_digit
-t --text_idle              The reverse of -i - IDLE events will be recorded as text "Idle"
-q --quiet                  Control amount of information displayed on console during processing:
                            -q      - no page number indications
                            -qq     - no page numbers or inflight analysis processing indications

-------------------------------------------------------------------------------------------------------------------------------


                            ===================
							Maintenance History
                            ===================
							
March 2023	GJN	Initial Creation
2023/03/23  GJN Added Throttle position translation for text fields
                Use RE to solve issue of 2 character TP values stuffing up line splitting.
2023/03/26  GJN Streamline worksheet writing code.
                Hide unwanted columns
                Protect code is written but not implemented as I'm still working out
                how to allow users to filter protected data...
                Add progress bar.

2023/03/27  GJN Stop using PDFReader as it was munging the data in cases where the TMC exceeded 3 digits
                or the Throttle Position was 2 characters.
                Now we use pdfplumber instead to extract the PDF data

2023/06/26  GJN For each annotation detected, print the time of the previous event and
                the timestamp delta. This will help when diagnosing the power up events.

2023/07/26  GJN When printing annotation events on event page, only print inter-event
                details for Power related events.

2023/08/31  GJN When writing to the Excel sheet, for the digital inputs, use a range function
                rather than individual writes.

2024/01/11  GJN Add ability to apply timeof day offset to logger records to compensate for
                the loss of accuracy of the logger's RTC.
                Offset is recorded in the configuration file.
                Compensation and other run time modifiers are recorded in 3rd worksheet
                as a matter of record

2024/07/11  GJN The data logger TOD clock has reverted to 1990 (epoch time) leading to problems
                with the record timestamps. A flag in the config file (epoch_timestamps_allowed)
                will control whether records with an epoch timestamp is allowed to be processed.

2024/07/12  GJN Refactor code to expand various variable names for clarity

2024/07/13  GJN FORKED OFF A VERSION TO READ TEXT FILES ISSUED USING THE GENERIC/TEXT
                PRINT DRIVER ON WINDOWS 11. PARSING IS A BIT MORE COMPLICATED BUT THE PDF
                PRINTING OPTION ON W11 IS NOT PRINTING ALL THE LINES REQUESTED - IT SEEMS
                TO STOP POPULATING THE FILE AFTER AROUND 1200 PAGES AND THEN GENERATES THE
                REMAINING PAGES WITH NO CONTENT.

2024/07/16  GJN First cut of text based input processing code

2024/07/17  GJN Add "writing_records_to_xls" flag to control writing 1990 date records to the
                workbook when we are filtering based on dates. We only want 1990 records included
                if they are within the range of selected dates.

2025/02/13  GJN In order to allow the code to cater for multiple locomotives, the wheel diameter actual
                setting in the configuration file is now a dictionary keyed on the locomotive number
                expressed as a string (to cater for odd bod entries). The locomotive number is extracted
                from the input file (so it must be read before the wheel diameter entry).
                If no locomotive number is detected in the input file or one is detected but there is no
                corresponding entry in the configuration file dictionary then the code will stop.
                The user should either fix the configuration file and/or manually edit the input file to
                set the "Locomotive Number is         -      xxx" line,

2025/02/25  GJN Add code to do rudimentary data analysis as records are processed. The initial cause for this
                is the requirement to identify anomalous traction motor current readings with the throttle in IDLE
                as part of a troubleshooting exercise. We found data samples where the throttle is in the IDLE position
                but which contain non-zero TMC readings, in some cases very high readings. There is a lag between
                commanding the idle position and the TMC decaying under the governor control, there's also the
                possibility of contactor arcing.

                We implement a fixed length queue for events, to retain the last (n - defined in the cfg file) events.
                When we hit an IDLE sample with a non-zero TMC value (or optionally, over a certain threshold) we will then log those
                events, plus the events leading up to the event in a separate sheet in the workbook.

                We continue to log readings until either the TMC drops to zero or the throttle
                goes out of idle. T

                In addition, we highlight epoch year cells in the primary worksheet to draw attention to them

2025/03/04  GJN hide unwanted columns in the in flight analysis worksheet. Set protection on all sheets

2025/03/06  GJN Add flag count sanity check to ensure the input file was printed with the correct (all) number of
                fields.

2025/7/20   GJN Add code to optionally suppress events where the locomotive is stationary for a certain number of events
                and there is no traction motor current
                Fix bug in write_annotations where row count was only incremented for a power related event in the logger 
                events sheet, thus overwriting other events such as laptop connects - these still showed up in the events
                sheet though.

2025/07/23  GJN Update stationary loco detection code to add Throttle Position in ID (Idle) as a third criteria for an
                idle locomotive

2025/07/25  GJN When suppressing events for stationary idle locomotive, instead of not writing the records, we now
                write all records but hide the suppressed records so the user can optionally  un-hide them later after
                providing the worksheet password.

2025/07/27  GJN Tidy up global var definition
                Add annotation when automatic brake pipe pressure transitions from 0 to non-zero state (compressor start)
                and the reverse (emergency brake application or train pipe break)
                The transition from non-zero BP pressure is a reasonable indication of engine start up (unless both
                brake stands are isolated for rotary servicing) which might help with correlation to the engine start
                time recorded in the loco log book in the event the logger RTC shows and epoch TOD state.

2025/07/29  GJN Allow brake pressures to be reported in kPa by setting a configuration file switch. Default is
                to report in psi.

2025/08/11  GJN Add command line argument code - to over-ride settings in configuration file.

2025/11/01  GJN Allow Throttle Position (Idle) to be recorded as either "Idle" for legibility or as 0 for
                chart generation ability. Controlled by configuration switch "idle_as_digit"

2025/11/02  GJN Add -i and -t switches to control Idle throttle notation from command line, change previous -t
                switch for timestamp adjustment to become -a to free up -t for this.
                Add -q (or -qq, -qqq etc.) switch and configuration item to control amount of processing reporting
                -q suppresses page numbers
                -qq suppresses page numbers and inflight analysis event indications on console

-------------------------------------------------------------------------------------------------------------------------------


NB: Low Idle position allows the engine to idle lower than normal to save fuel. 
	Not used on our 830 or 930 class locomotives.

	DYN refers to dynamic braking, the 830 class does not have this. The instances
	seen in the data were caused by spurious voltages appearing on the dynamic brake
	signal leads connected to the MU cable in 844's electrical cabinet. The presence
	of +74VDC on these wires masked the correct throttle settings. (Refer to the TP
	truth table in the Quantum Speedometer manual). We disconnected these 2 signal feeds
	into the data logger in June/July 2023 to fix this problem


-------------------------------------------------------------------------------------------------------------------------------


NOTE RE EPOCH OR 1990 DATE RECORD HANDLING

Due to an as yet unsolved, issue with the Quantum datalogger the time-of-day clock reset to
01/01/1990 in late April. The TOD counter incremented for a few minutes then reset and repeated this
until I reset the TOD in the logger in June.

This may well happen again so this code needs to be able to handle records with a TOD in the 1990
range (the Quantum system's epoch date). There is a flag in the configuration file to permit or
deny writing epoch records to the spreadsheet.

Also, when filtering records to those between 2 date/times - we need to include epoch records
within those 2 bounds but exclude epoch records prior to and after to the bounds. Note
that epoch dated records trailing a block of "in-date" records will be recorded until we reach
a non-epoch, out-of-date, record - this is because we cannot tell, when reading an epoch record,
whether the next non-epoch record somewhere further in the file is within or without the desired
range of dates.

Finally, when calculating the time between annotation records (non data records), we cannot calculate
the interval if either record has an epoch datestamp.

-------------------------------------------------------------------------------------------------------------------------------


"""

#import pprint
import os
import sys
import xlsxwriter
import argparse
from collections import deque
from datetime import datetime
import quantum_extraction_cfg as cfg


loco_number = ""
old_page_number=0
current_page_number=0
old_record_date="None"
old_record_time="None"
writing_records_to_xls=True     # Only used when filtering records based on date.

first_datestamp_written=[None,None]
last_non_epoch_datestamp_written=[None,None]
last_datestamp_written=[None,None]

previous_event_speed=-1 # Used to track loco stationary event sequences
previous_event_tmc=-1
previous_throttle_position=""
suppressed_stationary_event_count=0
first_suppressed_timestamp=""
last_suppressed_timestamp=""
suppressed_rows=list()      # Stores row numbers with suppressed events, used to hide said rows
previous_event_brake_pipe_pressure=-1   # Brake pipe pressure

global wb_name
global workbook
global ws_data_samples
global count_data_samples
global count_in_flight_analysis
global count_epoch_events
global count_suppressed_events
global ws_row_modifiers
global start_timestamp_epoch_seconds
global end_timestamp_epoch_seconds
global wheel_diameter_qdp_inches
global ws_row_data_samples
global ws_annotations
global ws_row_annotations
global ws_modifiers
global ws_in_flight_analysis
global ws_row_in_flight_analysis
global lalign
global cell_fill
global old_record_data
global in_suppression_mode



if cfg.in_flight_analysis_enabled:
    previous_events_deque = deque(maxlen = cfg.ifa_deque_maxlen)    # This will hold (n) data points for analysis

def main():
    # pp = pprint.PrettyPrinter(indent=4)

    global wb_name
    global workbook
    global ws_data_samples

    global first_datestamp_written
    global last_non_epoch_datestamp_written
    global last_datestamp_written

    global count_data_samples
    global count_in_flight_analysis
    global count_epoch_events
    global count_suppressed_events

    global ws_row_modifiers

    global start_timestamp_epoch_seconds
    global end_timestamp_epoch_seconds

    process_command_line_args()

    start_timestamp_epoch_seconds = get_epoch(cfg.start_timestamp)
    end_timestamp_epoch_seconds = get_epoch(cfg.end_timestamp)

    count_data_samples=0
    count_epoch_events=0
    count_suppressed_events=0
    if cfg.in_flight_analysis_enabled:
        count_in_flight_analysis=0

    if cfg.filter_dates:
        print("Record filtering enabled")
        print("Start from " + cfg.start_timestamp)
        print("End at " + cfg.end_timestamp)
    else:
        print("No record filtering required")
    if cfg.epoch_timestamps_allowed:
        print("Epoch year records are permitted")
        print("Epoch year is " + str(cfg.epoch_year))
    else:
        print("Epoch year records will be dropped")
    if cfg.ts_adjustment != 0:
        print("Timestamps adjustment factor is " + str(cfg.ts_adjustment) + " seconds")
    else:
        print("No timestamp adjustment in force")
    if cfg.suppress_stationary_events:
        print("Stationary loco events will be suppressed")
    else:
        print("Stationary loco events are included in report")
    if cfg.report_kpa_pressures:
        print("Pressures will be reported in kpa")

    print("Input = " + cfg.source_file)

    try:
        with open(cfg.source_file) as file:
            while raw_line := file.readline():
                # We need to examine the line to see if there is a FORM FEED (0x0C) within it, if so
                # the line needs to be split on that character and each half treated as a separate line
                # The W11 print to Generic Text or the Quantum software inserts FFs at the end of the page
                raw_line = raw_line.rstrip()
                lines = raw_line.split('\x0c')
                for line in lines:
                    process_line(line)
    except FileNotFoundError:
        print('Error: The file ',cfg.source_file, 'was not found.')
        sys.exit(-1)
    except PermissionError:
        print('Error: You do not have permission to access ', cfg.source_file)
        sys.exit(-1)
    except IOError as e:
        print(f"An I/O error occurred: {e}")
        sys.exit(-1)

    if cfg.suppress_stationary_events:
        hide_suppressed_rows(ws_data_samples,suppressed_rows)


    print("\nProcessing statistics")
    print("=====================")
    print(str(count_data_samples)+" data points processed")
    print(str(count_epoch_events)+" epoch dated events processed")
    if cfg.in_flight_analysis_enabled:
        print(str(count_in_flight_analysis)+" analysis streams processed")
    print(str(count_suppressed_events) + " stationary loco events suppressed")
    print("")
    print("First record written = "+first_datestamp_written[0]+" "+first_datestamp_written[1])
    print("Last record written =  "+last_datestamp_written[0]+" "+last_datestamp_written[1])
    if (last_non_epoch_datestamp_written[0] != last_datestamp_written [0]) and \
        (last_non_epoch_datestamp_written[1] != last_datestamp_written[1]):
        print("Last non-epoch record written =  " + last_non_epoch_datestamp_written[0] + " " + last_non_epoch_datestamp_written[1])


    ws_row_modifiers+=1
    ws_modifiers.write(ws_row_modifiers, 0, "Totals: "+str(count_data_samples)+" data points processed")
    ws_row_modifiers += 1
    ws_modifiers.write(ws_row_modifiers, 0, "Totals: "+str(count_epoch_events)+" epoch dated events processed")
    ws_row_modifiers += 1
    if cfg.in_flight_analysis_enabled:
        ws_modifiers.write(ws_row_modifiers, 0,"Totals: " + str(count_in_flight_analysis)+" analysis streams processed")
        ws_row_modifiers += 1
    ws_modifiers.write(ws_row_modifiers, 0, "Totals: "+str(count_suppressed_events)+" stationary loco events suppressed")
    ws_row_modifiers += 1


    hide_columns(ws_data_samples, cfg.headers)
    if cfg.in_flight_analysis_enabled:
        hide_columns(ws_in_flight_analysis, cfg.headers)
    ws_data_samples.protect(cfg.protect_string,cfg.protection_mode)
    ws_annotations.protect(cfg.protect_string,cfg.protection_mode)
    ws_modifiers.protect(cfg.protect_string,cfg.protection_mode)
    if cfg.in_flight_analysis_enabled:
        ws_in_flight_analysis.protect(cfg.protect_string, cfg.protection_mode)
    workbook.close()
    print("Written file : " + wb_name)



def process_line(line):
    """
        Process each line, if we are in page 1 we set a number of variables based on the contents.
        For other pages, if the line starts with a number (ie a date record) then we pass it to the data sampling function
        otherwise it's an annotation so we write it to the annotation worksheet
    """
    global old_page_number
    global loco_number
    global wheel_diameter_qdp_inches
    global current_page_number

    if len(line)==0:
        return

    if 'Page' in line:
        current_page_number = get_page_number(line)
        if cfg.quiet==0:
            print("Processing page " + str(current_page_number))
        return

    if old_page_number > 1:
        # Skip lines with strings we are not interested in
        if skip_line_found(line):
            return

        # Data lines begin with an integer (1st character in timestamp)
        if line[0].isnumeric():  # Data sample lines are the only ones starting with a digit
            process_sample(line)
        else:
            write_annotation(line,True)

        if current_page_number != old_page_number:
            old_page_number=current_page_number

    else:
        # Page 1 stuff here
        # If we are now on page number 2 then we should have all the informational variables from
        # page 1 set and ready to create the workbook.
        if current_page_number == 2 and old_page_number == 1:
            create_workbook()
            old_page_number=current_page_number
            return
        old_page_number=current_page_number
        if "Locomotive Number" in line:
            words = line.split()
            loco_number = words[-1]
            return
        # The wheel diameter adjustment factor is based on the wheel diameter reported from the input file
        # combined with the actual wheel diameter defined in the configuration file. Because this code caters
        # for multiple locomotives there will be a configuration entry for each, in a dictionary keyed by the
        # loco number, obtained from the input file. If there is no loco number in the input file or there is no
        # match in the configuration file then the code will stop.
        if cfg.speed_adjustment_factor == 0:
            if "Circumference" in line and "Diameter" in line:
                words = line.split()
                # pp.pprint(words)
                # Check for wheel size entry in config dictionary
                if loco_number == "":   # not set
                    print("No locomotive number detected in the input file. Please check and set.")
                    sys.exit(1)
                if loco_number in cfg.wheel_dia_actual_mm:
                    actual_wheel_dia_mm = cfg.wheel_dia_actual_mm[loco_number]
                else:
                    print("No wheel diameter defined in configuration file for locomotive "+loco_number)
                    sys.exit(1)
                wheel_diameter_qdp_inches = float(words[-1])  # wheel diameter according to the QDP software
                cfg.speed_adjustment_factor = actual_wheel_dia_mm / (wheel_diameter_qdp_inches * 25.4)
                # pp.pprint(cfg.speed_adjustment_factor)
                return
        return  # We don't want anything else from page 1

def create_workbook():
    """
        Create Excel workbook with required pages and initiate vars for each page to track the current row for that page
    """
    global loco_number
    global workbook
    global ws_data_samples
    global ws_row_data_samples
    global ws_annotations
    global ws_row_annotations
    global ws_modifiers
    global ws_row_modifiers

    if cfg.in_flight_analysis_enabled:
        global ws_in_flight_analysis
        global ws_row_in_flight_analysis

    global lalign
    global cell_fill
    global wb_name
    global wheel_diameter_qdp_inches


    wb_name = cfg.workbook_name + " " + loco_number + " " + datetime.now().strftime("%Y%m%d%H%M") + ".xlsx"
    workbook = xlsxwriter.Workbook(wb_name, {'strings_to_numbers': True})
    ws_data_samples = workbook.add_worksheet(cfg.worksheet_name)
    lalign = workbook.add_format({'align': 'left'})
    cell_fill = workbook.add_format({'bg_color': 'yellow'})
    ws_annotations = workbook.add_worksheet("Logger Events")
    parts = os.path.split(cfg.source_file)
    ws_modifiers = workbook.add_worksheet("Runtime modifiers")
    ws_row_data_samples = write_header(workbook, ws_data_samples, "Data extract from Quantum Data Recorder",
                            "Locomotive " + loco_number + ". Source file " + parts[1])
    ws_row_annotations = write_header_ann(workbook, ws_annotations,
                        "Data extract from Quantum Data Recorder", loco_number)
    ws_row_modifiers = write_header_modifiers(workbook, ws_modifiers, "Runtime modifiers and events")
    if cfg.filter_dates:
        ws_modifiers.write(ws_row_modifiers, 0,
                            "Records selected from " + cfg.start_timestamp + " to " + cfg.end_timestamp)
        ws_row_modifiers += 1
    else:
        ws_modifiers.write(ws_row_modifiers, 0, "No record filtering in place")
        ws_row_modifiers += 1
    ws_modifiers.write(ws_row_modifiers, 0,
                        "Record timestamp offset applied is " + str(cfg.ts_adjustment) + " seconds")
    ws_row_modifiers += 1
    ws_modifiers.write(ws_row_modifiers, 0,
                       "Speed adjustment factor applied. QDP defined wheel diameter = " + str(
                        wheel_diameter_qdp_inches) + " inches (" + str(
                       wheel_diameter_qdp_inches * 25.4) + " mm). Measured wheel diameter = " + str(
                       cfg.wheel_dia_actual_mm[loco_number]) + " mm. Adjustment factor = " + str(
                       cfg.speed_adjustment_factor) + ".")
    ws_row_modifiers += 1
    if cfg.epoch_timestamps_allowed:
        ws_modifiers.write(ws_row_modifiers, 0,
            "Epoch dated records permitted. Epoch year is " + str(cfg.epoch_year))
        ws_row_modifiers += 1
    else:
        ws_modifiers.write(ws_row_modifiers, 0, "Epoch year (" + str(cfg.epoch_year) + ") dated records omitted")
        ws_row_modifiers += 1

    if cfg.in_flight_analysis_enabled:
        ws_in_flight_analysis = workbook.add_worksheet("Event Analysis")
        ws_row_in_flight_analysis = write_header(workbook, ws_in_flight_analysis, "Event of interest analysis",
                                           "Locomotive " + loco_number + ". Source file " + parts[1])
        ws_in_flight_analysis.write(ws_row_in_flight_analysis,0,"Events will be flagged if the TMC value is over "+str(cfg.ifa_tmc_threshold)+" Amps with the throttle in IDLE")
        ws_row_in_flight_analysis+=1
        ws_in_flight_analysis.write(ws_row_in_flight_analysis,1,"This may be caused by arcing across contactors when dropping to Idle position.")
        ws_row_in_flight_analysis+=1
        ws_in_flight_analysis.write(ws_row_in_flight_analysis,0,"The previous "+str(cfg.ifa_deque_maxlen)+" events will be shown. All subsequent events will also be shown until the selection criteria are no longer met")
        ws_row_in_flight_analysis+=2
        ws_modifiers.write(ws_row_modifiers,0,"Event analysis: Events will be flagged if the TMC value is over "+str(cfg.ifa_tmc_threshold)+" Amps with the throttle in IDLE")
        ws_row_modifiers+=1
        ws_modifiers.write(ws_row_modifiers,0,"Event analysis: The previous "+str(cfg.ifa_deque_maxlen)+" events will be shown. All subsequent events will also be shown until the selection criteria are no longer met")
        ws_row_modifiers+=1

    if cfg.suppress_stationary_events:
        ws_modifiers.write(ws_row_modifiers, 0,
                           "Events where locomotive is stationary (speed = 0 kph, throttle is in idle, and tmc = 0) are suppressed.")
        ws_row_modifiers += 1

    if cfg.report_kpa_pressures:
        ws_modifiers.write(ws_row_modifiers, 0,
                           "Brake system pressures reported in kpa.")
        ws_row_modifiers += 1

    return

def write_annotation(line,write_to_logger_event_sheet):
    """
        Annotations are text records that contain no loco movement data, they get written to a worksheet in the workbook.
        The code will also add records to this worksheet to record activities of interest
    """
    global start_timestamp_epoch_seconds
    global end_timestamp_epoch_seconds
    global ws_data_samples
    global ws_row_data_samples
    global ws_annotations
    global ws_row_annotations
    global old_record_data
    global old_record_time


    # Handle annotations
    # Extract date and time - last word in string in format HH:MM:SS-mm/dd/yyyy
    words = line.split()
    record_date = convert_date(words[-1])
    record_time = words[-2].replace("-", "")
    record_date, record_time = apply_time_adjustment(record_date, record_time)
    record_ts_epoch_seconds = get_epoch(record_date + " " + record_time)
    if start_timestamp_epoch_seconds > 0 and (
            (record_ts_epoch_seconds < start_timestamp_epoch_seconds) or (
            record_ts_epoch_seconds > end_timestamp_epoch_seconds)):
            return

    ws_data_samples.write(ws_row_data_samples, 0, record_date)
    ws_data_samples.write(ws_row_data_samples, 1, record_time)
    ws_data_samples.write(ws_row_data_samples, 2, ' '.join(words[:-2]), lalign)
    ws_row_data_samples += 1

    if not write_to_logger_event_sheet:  # only write to data samples sheet.
        return


    # Calculate offset between this annotation and the previous record.
    # If either date is in the epoch period then don't do this as it makes no sense
    if old_record_date != "None" and not check_for_epoch_year(old_record_date) and not check_for_epoch_year(record_date):
        s = datetime.strptime(old_record_date + " " + old_record_time, "%Y/%m/%d %H:%M:%S")
        e = datetime.strptime(record_date + " " + record_time, "%Y/%m/%d %H:%M:%S")
        offset = str(e - s)
    else:
        offset="N/A"

    ws_annotations.write(ws_row_annotations, 0, record_date)
    ws_annotations.write(ws_row_annotations, 1, record_time)
    ws_annotations.write(ws_row_annotations, 2, " ".join(words[:-2]))
    # Only write inter-event interval for power related events.
    if words[0][:5] == 'Power':
        ws_annotations.write(ws_row_annotations, 3, old_record_date)
        ws_annotations.write(ws_row_annotations, 4, old_record_time)
        ws_annotations.write(ws_row_annotations, 5, offset)
    ws_row_annotations += 1

    return

def process_sample(line):
    """
        This function is passed a line containing data from the Quantum data logger, the function parses the data and
        passes it to be written to the Excel worksheet
    """
    global old_record_date
    global old_record_time
    global ws_data_samples
    global ws_row_data_samples
    global count_data_samples
    global count_epoch_events
    global writing_records_to_xls           # Only used when filtering records on date
    global first_datestamp_written
    global last_non_epoch_datestamp_written
    global last_datestamp_written
    global previous_event_speed
    global previous_event_tmc
    global suppressed_stationary_event_count
    global in_suppression_mode
    global count_suppressed_events
    global first_suppressed_timestamp
    global last_suppressed_timestamp
    global previous_throttle_position
    global suppressed_rows
    global previous_event_brake_pipe_pressure

    time_position = 0
    time_length = 8
    date_position = 10
    date_length = 10
    remainder_position = 20
    speed_position = 28
    speed_length = 4
    tmc_position = 32
    tmc_length = 4

    # Date and time are in fixed positions starting at column 0 and in the format
    # hh:mm:ss- mm/dd/yyyy
    record_time = line[time_position:time_position + time_length]
    date_us_fmt = line[date_position:date_position + date_length]
    record_date = convert_date(date_us_fmt)
    record_date, record_time = apply_time_adjustment(record_date, record_time)
    old_record_date = record_date
    old_record_time = record_time

    # Because the mileage field leading space is lost when the distance goes to 3 figures and
    # extends when it goes to 4 figures, we start at the end of the date field then strip any leading
    # spaces. The mileage field should be followed by a space as the next field will be the speed
    parts = line[remainder_position:].lstrip().split(" ", 1)
    mileage = float(parts[0])
     # We need to handle speed and tmc carefully as the TMC field will lose leading spaces when it
    # goes over 3 digits
    # pp.pprint(line[28:32])
    speed = int(line[speed_position:speed_position + speed_length].strip())
    tmc = int(line[tmc_position:tmc_position + tmc_length].strip())
    parts = line[speed_position + speed_length + tmc_length:].lstrip().split()
    # pp.pprint(parts)

    # The first 3 fields are values as follows:
    brake_pipe_pressure = int(parts[0])
    if cfg.report_kpa_pressures:
        brake_pipe_pressure=round(brake_pipe_pressure*cfg.psi_to_kpa_factor)
    brake_cylinder_pressure = int(parts[1])
    if cfg.report_kpa_pressures:
        brake_cylinder_pressure=round(brake_cylinder_pressure*cfg.psi_to_kpa_factor)

    throttle_position = parts[2]  # This is left as string to cater for (D)dynamic or low (ID)le states

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
    flags = parts[3:]
    # Sanity check the flags to ensure we have the correct number. If the print setup is
    # wrong in the QDP software then this may occur. If this were allowed to go through then
    # the column headers for the flags would be wrong!
    if len(flags) != cfg.number_of_flags_expected:
        print("FATAL: Expected "+str(cfg.number_of_flags_expected)+" flags but received "+str(len(flags))+" in line ["+line+"]. Processing abandoned")
        sys.exit(1)


    record_ts_epoch_seconds = get_epoch(record_date + " " + record_time)
    # We want to filter out dates prior to or after a range of datestamps - use timestamp WITHOUT adjustments
    is_epoch_year_datestamp = check_for_epoch_year(record_date)
    if is_epoch_year_datestamp:
        fill = True
        count_epoch_events+=1
    else:
        fill = False
    if cfg.filter_dates:
        # Timestamp is epoch year and epoch year timestamps are not allowed then skip the write step
        if is_epoch_year_datestamp and not cfg.epoch_timestamps_allowed:
            return
        # Timestamp is epoch year, but we are not currently writing records to the workbook as they
        # are outside the required date range
        if is_epoch_year_datestamp and not writing_records_to_xls:
            return
        # Timestamp is NOT an epoch year and record timestamp is outside desired range then skip the write step
        # Set flag to prevent epoch year records from being written to the workbook.
        if not is_epoch_year_datestamp and ((record_ts_epoch_seconds < start_timestamp_epoch_seconds) or (
                    record_ts_epoch_seconds > end_timestamp_epoch_seconds)):
            writing_records_to_xls=False
            return
        # Set a flag to allow epoch year records to be written the workbook as we are in the range of
        # valid date records.
        writing_records_to_xls=True

    # Write record to spreadsheet
    if first_datestamp_written[0] is None:
        first_datestamp_written[0]=record_date
        first_datestamp_written[1]=record_time

    # Check for brake pipe pressure changes of interest
    #       Transition from 0 to non-zero - engine startup?
    if cfg.report_kpa_pressures:
        pressure_unit="kpa"
    else:
        pressure_unit="psi"
    if previous_event_brake_pipe_pressure == 0 and brake_pipe_pressure > 0:     # Compressor start up
        write_annotation("Brake pipe pressure transitioned from "+str(previous_event_brake_pipe_pressure)+" "+pressure_unit+" to "+str(brake_pipe_pressure)+" "+pressure_unit+" - compressor start up "+line[time_position:remainder_position],True)
    #       Transition from non-zero tp 0 - emergency application or brake pipe rupture?
    if previous_event_brake_pipe_pressure > 0  and brake_pipe_pressure == 0:
        write_annotation("Brake pipe pressure transitioned from "+str(previous_event_brake_pipe_pressure)+" "+pressure_unit+" to "+str(brake_pipe_pressure)+" "+pressure_unit+". "+line[time_position:remainder_position],True)

    ##################################################################################################
    # NOTE: Any state change that writes an annotation to the data samples sheet MUST be done prior  #
    #       to the following line suppression code otherwise the annotation row will be added to the #
    #       list of rows to be hidden instead of the actual data sample row number!                  #
    ##################################################################################################

    # speed is zero, previous speed was zero, TP - ID(le) and we are suppressing stationary events
    # don't write this record to the sheet. increment the counter
    if cfg.suppress_stationary_events and speed==0 and previous_event_speed==0 and tmc==0 and previous_event_tmc==0 and throttle_position=="ID" and previous_throttle_position=="ID":
        if suppressed_stationary_event_count==0:
            first_suppressed_timestamp=record_time
        suppressed_stationary_event_count+=1
        last_suppressed_timestamp=record_time
        suppressed_rows.append(ws_row_data_samples)   # These will be hidden in due course


    # speed is not zero, previous speed was zero, and we are suppressing stationary events
    # write this record to the sheet after reporting the gap in events...
    if cfg.suppress_stationary_events and (speed!=0 and previous_event_speed==0) or (tmc!=0 and previous_event_tmc==0) or (throttle_position!="ID" and previous_throttle_position=="ID"):
        if suppressed_stationary_event_count!=0:
            write_annotation("Suppressed "+str(suppressed_stationary_event_count)+" consecutive "+("event" if suppressed_stationary_event_count==1 else "events")+" with Speed = 0 kph, TMC = 0 Amps, and Throttle in Idle from "+first_suppressed_timestamp+" to "+last_suppressed_timestamp+" "+line[time_position:remainder_position],False)
            count_suppressed_events+=suppressed_stationary_event_count
        suppressed_stationary_event_count=0





    ws_row_data_samples = write_record(ws_data_samples,
                                       ws_row_data_samples,
                                       mileage,
                                       speed,
                                       tmc,
                                       brake_pipe_pressure,
                                       brake_cylinder_pressure,
                                       throttle_position,
                                       flags,
                                       record_date,
                                       record_time,
                                       fill,
                                       False)
    previous_event_speed=speed
    previous_event_tmc=tmc
    previous_throttle_position=throttle_position
    previous_event_brake_pipe_pressure=brake_pipe_pressure

    if not is_epoch_year_datestamp:
        last_non_epoch_datestamp_written[0] = record_date
        last_non_epoch_datestamp_written[1] = record_time
    last_datestamp_written[0]=record_date
    last_datestamp_written[1]=record_time

    # If we are doing in flight analysis, add this record to the deque.
    # Later we'll play more with this stuff.
    if cfg.in_flight_analysis_enabled:
        previous_events_deque.append((record_date,record_time,mileage,speed,tmc,brake_pipe_pressure,brake_cylinder_pressure,throttle_position,flags))
        perform_in_flight_analysis()

    count_data_samples+=1

    return

def perform_in_flight_analysis():
    """
        Analyse data records and alert on events of interest. A deque is set up to store the 10 events leading up to the current event so
        we can report on precursors to an event
    """


    global ws_in_flight_analysis
    global ws_row_in_flight_analysis
    global count_in_flight_analysis
    # At this stage the current data point has been appended to the deque, so we have up to (n-1) previous data points plus the current data point
    # each data point is a tuple with the following members:
    # 0 - date              5 - bp pressure
    # 1 - time              6 - bc pressure
    # 2 - mileage           7 - throttle position (1-8, ID, LO etc.)
    # 3 - speed             8 - binary flags (11 off)
    # 4 - tmc

    # See if this data point contains an item of interest
    current_data_point = previous_events_deque[-1]
    # If we are in an event of interest we need to write this data point to the output file irrespective of
    # its contents then check whether we reset the event of interest flag
    if cfg.ifa_in_event_of_interest:
        #print(current_data_point)
        if current_data_point[7] != 'ID' or current_data_point[4] == 0:
            fill=False
        else:
            fill=True
        ws_row_in_flight_analysis = write_record(ws_in_flight_analysis,
                                           ws_row_in_flight_analysis,
                                           current_data_point[2],
                                           current_data_point[3],
                                           current_data_point[4],
                                           current_data_point[5],
                                           current_data_point[6],
                                           current_data_point[7],
                                           current_data_point[8],
                                           current_data_point[0],
                                           current_data_point[1],
                                         False,
                                                 fill)
        # If the throttle leaves the idle position or the tmc drops to 0 then we are done with this event
        if current_data_point[7] != 'ID' or current_data_point[4] == 0:
            cfg.ifa_in_event_of_interest = False
            ws_in_flight_analysis.write(ws_row_in_flight_analysis, 0, "End of event flow "+str(count_in_flight_analysis))
            ws_row_in_flight_analysis += 2
            if cfg.quiet < 2:
                print("EVENT "+str(count_in_flight_analysis)+" TERMINATED")
        return


    # Check throttle position - we are only interested in Idle state
    if current_data_point[7] != 'ID':
        return
    # Now check the tmc value. If the threshold is set to > 0 then we are only interested in values greater than
    # or equal to the threshold. If the threshold is zero then we want all non-zero IDLE events
    if cfg.ifa_tmc_threshold != 0 and current_data_point[4] < cfg.ifa_tmc_threshold:
        return
    if cfg.ifa_tmc_threshold == 0 and current_data_point[4] == 0:
        return

    # We are now in an event of interest so we log all the events in the deque and set the flag
    count_in_flight_analysis+=1
    if cfg.quiet < 2:
        print("EVENT "+str(count_in_flight_analysis)+" COMMENCED")
    ws_in_flight_analysis.write(ws_row_in_flight_analysis, 0, "Start of event flow "+str(count_in_flight_analysis))
    ws_row_in_flight_analysis += 1

    deque_len=len(previous_events_deque)
    for index, preceding_data_point in enumerate(previous_events_deque):
        #print(preceding_data_point)
        if index == deque_len-1:
            fill=True
        else:
            fill=False
        ws_row_in_flight_analysis = write_record(ws_in_flight_analysis,
                                           ws_row_in_flight_analysis,
                                           preceding_data_point[2],
                                           preceding_data_point[3],
                                           preceding_data_point[4],
                                           preceding_data_point[5],
                                           preceding_data_point[6],
                                           preceding_data_point[7],
                                           preceding_data_point[8],
                                           preceding_data_point[0],
                                           preceding_data_point[1],
                                    False,
                                                 fill)

    cfg.ifa_in_event_of_interest = True


    return

def hide_columns(ws, headers):
    """ Hide any column with False in the header tuple """
    for column, record in enumerate(headers):
        if not record[1]:
            ws.set_column(column, column, None, None, {'hidden': True})


def write_record(ws, ws_row, mileage, speed, tmc, brake_pipe_pressure, brake_cylinder_pressure, throttle_position, flags, record_date, record_time, fill_year_cell,fill_tmc_cell):
    """ Write spreadsheet row, return updated row number """

    # Time - has a dash appended
    # Date - is in mm/dd/yyyy format
    # Miles - 2 decimal places
    # Speed - MPH?
    # Traction motor current
    # Brake pipe pressure
    # Independent brake pressure
    # Throttle notch

    # Flags are either 1 or 0 for ON or OFF
    # Reverser in reverse
    # Engineer induced emergency
    # Pressure control switch (set when the BP air drops below 45 psi)
    # Headlight on - short end
    # Reverser in forward
    # Headlight on - long end
    # Horn on
    # Digital spare 1
    # Digital spare 2
    # Vigilance Control Alert acknowledge
    # Axle drive type


    ws_col = 0
    if fill_year_cell:
        ws.write_string(ws_row, ws_col, record_date,cell_fill)  # AUS Date stamp
    else:
        ws.write_string(ws_row, ws_col, record_date)  # AUS Date stamp

    ws_col += 1
    ws.write_string(ws_row, ws_col, record_time)  # Timestamp
    ws_col += 1
    # ws.write(ws_row, ws_col, "{:.2f}".format(float(words[2]) * 1.6))  # Mileage converted to km units
    ws.write_number(ws_row, ws_col, float(mileage) * 1.6)  # Mileage converted to km units
    ws_col += 1
    # Speed, converted to kph and adjusted according to the difference between the real wheel diameter
    # and the diameter reported by the QDP software. NB: The reported wheel diameter can be set when
    # downloading the data via QDP but not when downloading via the QRST software.
    ws.write_number(ws_row, ws_col, round(speed * 1.6 * cfg.speed_adjustment_factor))
    ws_col += 1
    if fill_tmc_cell:
        ws.write_number(ws_row, ws_col, tmc, cell_fill)  # TMC
    else:
        ws.write_number(ws_row, ws_col, tmc)  # TMC
    ws_col += 1
    ws.write_number(ws_row, ws_col, brake_pipe_pressure)  # Brake pipe pressure
    ws_col += 1
    ws.write_number(ws_row, ws_col, brake_cylinder_pressure)  # Independent brake pressure
    ws_col += 1
    ws.write(ws_row, ws_col, translate_tp(throttle_position))  # Throttle position
    ws_col += 1
    # Digital inputs follow
    for flag in flags:
        ws.write(ws_row, ws_col, "Y" if flag == "1" else "N")
        ws_col += 1

    ws_row += 1
    return ws_row


def translate_tp(tp):
    """ take a throttle position. If it's a number, then return that number.
        If it's a letter then returnn the corresponding text.
        If the TP iw Idle and the control flag is set then we return 0 rather then Idle """
    if tp.isnumeric():
        return tp
    if tp.upper()[0]=="I" and cfg.idle_as_digit:
        return 0
    if tp.upper() in cfg.tp_translations.keys():
        return cfg.tp_translations[tp.upper()]
    return tp + " (Unknown)"


def convert_date(us_date):
    """ Convert US formatted date to AUS formatted date """
    parts = us_date.split("/")
    if len(parts) != 3:
        return "invalid"
    aus_date = "{:0>4d}/{:0>2d}/{:0>2d}".format(int(parts[2]), int(parts[0]), int(parts[1]))
    return aus_date


def check_for_epoch_year(date):
    """
        Return true if the date contains the epoch year (usually 1990)
    """
    if str(cfg.epoch_year) in date:
      return True
    return False


def write_header(wb, ws, text, loco_number):
    """
        write the header(s) to an Excel worksheet
    """

    wb.set_size(1920, 1080)
    lalign = wb.add_format({'align': 'left'})
    ws.set_column('A:B', 15, lalign)
    numfmt = wb.add_format({'num_format': '0.00'})
    ws.set_column('C:C', 10, numfmt)
    ralign = wb.add_format({'align': 'right'})
    ws.set_column('D:H', 10, ralign)
    calign = wb.add_format({'align': 'center'})
    ws.set_column('I:S', 10, calign)
    ws.set_column('T:T', 20, lalign)

    calign_b = wb.add_format({'align': 'center', 'bold': True})
    ws.set_row(1, None, calign_b)

    header_format = wb.add_format({'font_size': 14, 'bold': True})

    ws.freeze_panes(3, 0)

    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0, 0, text + " : " + loco_number, header_format)

    if cfg.report_kpa_pressures:
        # Change header value from config file if we are reporting in kpa.
        pressure_unit="(kpa)"
    else:
        pressure_unit="(psi)"
    for column, record in enumerate(cfg.headers):
        ws.write(1, column, record[0].replace("(psi)",pressure_unit))
    return 3


def write_header_modifiers(wb, ws, text):
    """
        Write the header row for the modifiers worksheet
    """
    l_align = wb.add_format({'align': 'left'})
    ws.set_column('A:A', 150, l_align)
    header_format_modifiers = wb.add_format({'font_size': 14, 'bold': True})
    ws.freeze_panes(3, 0)
    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0, 0, text, header_format_modifiers)
    return 3


def write_header_ann(wb, ws, text, loco_number):
    """
        Write the header row for the annotations worksheet
    """
    l_align = wb.add_format({'align': 'left'})
    ws.set_column('A:B', 15, l_align)
    ws.set_column('C:C', 50, l_align)
    ws.set_column('D:F', 15, l_align)

    header_format_ann = wb.add_format({'font_size': 14, 'bold': True})
    ws.freeze_panes(3, 0)
    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0, 0, text + " : " + loco_number, header_format_ann)
    ws.write(1, 0, "Event Date", header_format_ann)
    ws.write(1, 1, "Event Time", header_format_ann)
    ws.write(1, 2, "Event Type", header_format_ann)
    ws.write(1, 3, "Prev Evt Date", header_format_ann)
    ws.write(1, 4, "Prev Evt Time", header_format_ann)
    ws.write(1, 5, "Offset", header_format_ann)
    return 3


def skip_line_found(line):
    """
        Search for existence of skip_list word(s) in the line variable passed into the function.
        If found then return True, otherwise return False
    """
    for word in cfg.skip_list_words:
        if word in line:
            return True
    return False


def get_epoch(timestamp):
    """ Take string in format yyyy/mm/dd hh:mm:ss and return epoch seconds (or 0 if flag is false) """
    if not cfg.filter_dates:
        return 0
    d = datetime.strptime(timestamp, "%Y/%m/%d %H:%M:%S")
    epoch = datetime(d.year, d.month, d.day, d.hour, d.minute, d.second).timestamp()
    return int(epoch)


def isfloat(num):
    """
        Test for a float value
    """
    try:
        float(num)
        return True
    except ValueError:
        return False


def apply_time_adjustment(date, time):
    """
        The logger realtime clock can vary from the actual time so the timestamps
        are not accurate. This function will normalise the timestamps based on the
        TOD adjustment factor. If the logger clock is behind the real clock then a
        positive adjustment is made, if ahead then a negative adjustment is made.
        The adjustment factor is in seconds.
    """
    d = datetime.strptime(date + " " + time, "%Y/%m/%d %H:%M:%S")
    epoch = datetime(d.year, d.month, d.day, d.hour, d.minute, d.second).timestamp()
    epoch += cfg.ts_adjustment
    datetime_obj = datetime.fromtimestamp(epoch)
    date = datetime_obj.strftime("%Y/%m/%d")
    time = datetime_obj.strftime("%H:%M:%S")
    return date, time

def get_page_number(line):
    """
        Find the page number from the incoming pages
    """
    parts = line.rstrip().split()
    page_number = int(parts[4])
    return page_number

def hide_suppressed_rows(ws,suppressed_rows):
    """
        Passed a worksheet and a list of row numbers, hide each of the rows from the list
    """
    if len(suppressed_rows)==0:
        return
    for row in suppressed_rows:
        ws.set_row(row,None,None,{'hidden':True})

def process_command_line_args():
    """
        Command line arguments may over-ride the directives in the config file
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--filename',
                        help='if set, this file path over-rides the entry in the configuration file')
    parser.add_argument('-a', '--ts_adjust',
                        help='if set, this value in seconds is applied to the logger clock timestamps to bring them in sync with the real time clock')
    parser.add_argument('-b', '--begin_timestamp',
                        help='if set, filters record by date - should be in the form yyyy/mm/dd hh:mm:ss - end timestamp must also be supplied')
    parser.add_argument('-e', '--end_timestamp',
                        help='if set, filters record by date - should be in the form yyyy/mm/dd hh:mm:ss - start timestamp must also be supplied')
    parser.add_argument('-k','--kpa_pressures', help='if set, pressures are reported in metric units', action='store_true' )
    parser.add_argument('-p','--psi_pressures', help='if set, pressures are reported in imperial units', action='store_true' )
    parser.add_argument('-s', '--suppress_stationary', help='if set, hide rows where loco is stationary, TMC=0 and TP is Idle',
                        action='store_true')
    parser.add_argument('-n', '--no_suppress_stationary', help='if set, show rows where loco is stationary, TMC=0 and TP is Idle',
                        action='store_true')
    parser.add_argument('-i','--integer_idle', help='if set, throttle position idle is reported as integer 0', action='store_true' )
    parser.add_argument('-t','--text_idle', help='if set, throttle position idle is reported as Idle', action='store_true' )
    parser.add_argument('-q','--quiet', action='count', default=0, help='Modify progress display on console. -q = no page numbers, -qq = no in-flight-analysis counts or page numbers, ')
    args = parser.parse_args()

    if args.filename:
        print("CFG source file ", cfg.source_file, " over-ridden by command line value ", args.filename)
        cfg.source_file = args.filename
    if args.ts_adjust:
        print("CFG timestamp adjustment ", cfg.ts_adjustment, " over-ridden by command line value ", args.ts_adjust)
        cfg.ts_adjustment = args.ts_adjust
    if args.begin_timestamp:
        if not args.end_timestamp:
            print("If supplying a start timestamp for record filtering, you must also supply an end timestamp")
            sys.exit(-1)
        print("CFG record filtering enabled. Start timestamp ", cfg.start_timestamp,
              " over-ridden by command line value ", args.begin_timestamp)
        cfg.start_timestamp = args.begin_timestamp
        cfg.filter_dates = True
    if args.end_timestamp:
        if not args.start_timestamp:
            print("If supplying an end timestamp for record filtering, you must also supply a start timestamp")
            sys.exit(-1)
        print("CFG record filtering enabled. End timestamp ", cfg.end_timestamp,
              " over-ridden by command line value ", args.end_timestamp)
        cfg.end_timestamp = args.end_timestamp
        cfg.filter_dates = True
    if args.kpa_pressures:
        print("CFG air pressures will be reported in kpa")
        cfg.report_kpa_pressures = True
    if args.psi_pressures:
        print("CFG air pressures will be reported in psi")
        cfg.report_kpa_pressures = False
    if args.suppress_stationary:
        print("CFG suppress stationary loco events")
        cfg.suppress_stationary_events = True
    if args.no_suppress_stationary:
        print("CFG do not suppress stationary loco events")
        cfg.suppress_stationary_events = False
    if args.integer_idle:
        print("CFG idle_as_digit over-ridden to report idle as integer")
        cfg.idle_as_digit = True
    if args.text_idle:
        print("CFG idle_as_digit over-ridden to report idle as text")
        cfg.idle_as_digit = False
    if args.quiet > 0:
        print("CFG quiet value of " + str(cfg.quiet) + " over-ridden by CLI switch value "+ str(args.quiet))
        cfg.quiet=args.quiet


if __name__ == '__main__':
    main()
