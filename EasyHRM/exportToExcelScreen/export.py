from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill
from itertools import chain

global EVENT_COLOR
EVENT_COLOR = "F0FC5A"
global disabled_sections
disabled_sections = []
global HIGH_AMPLITUDE_MINIMUM
HIGH_AMPLITUDE_MINIMUM_VALUE = 100
global HIGH_AMPLITUDE_LENGTH
HIGH_AMPLITUDE_MINIMUM_PATTERN_LENGTH = 3
global HIGH_AMPLITUDE_MINIMUM_LENGTH_CM
HIGH_AMPLITUDE_MINIMUM_LENGTH_MM = 100
LONG_PATTERN_MINIMUM_SENSORS = 5
ALTERNATING_EVENT_COLORS = ["F0FC5A", "FDE9D9"]

def initialize_comprehensive_statistics(event_names):
    """Initialize the comprehensive statistics structure for the new table"""
    regions = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    pattern_types = [
        "Long s", "Short s", "Long r", "Short r", "Long a", "Short a",
        "Long ", "Short "  # Handle cases where direction might be empty
    ]
    region_ranges = [
        "Ascending - Rectum", "Transverse - Rectum", 
        "Descending - Rectum", "Sigmoid-Rectum"
    ]
    
    comprehensive_stats = {}
    
    for event in event_names:
        comprehensive_stats[event] = {}
        
        # Initialize all pattern type combinations
        for pattern_type in pattern_types:
            comprehensive_stats[event][pattern_type] = {}
            
            # Regional stats
            for region in regions:
                comprehensive_stats[event][pattern_type][region] = {
                    'count': 0, 'velocities': [], 'amplitudes': []
                }
            
            # Region-range stats
            for region_range in region_ranges:
                comprehensive_stats[event][pattern_type][region_range] = {
                    'count': 0, 'velocities': [], 'amplitudes': []
                }
            
            # Totals
            comprehensive_stats[event][pattern_type]['Total'] = {
                'count': 0, 'velocities': [], 'amplitudes': []
            }
        
        # Special categories
        comprehensive_stats[event]['cyclic s'] = {'count': 0, 'velocities': [], 'amplitudes': []}
        comprehensive_stats[event]['cyclic r'] = {'count': 0, 'velocities': [], 'amplitudes': []}
        comprehensive_stats[event]['cyclic a'] = {'count': 0, 'velocities': [], 'amplitudes': []}
        comprehensive_stats[event]['HAPCs'] = {'count': 0, 'velocities': [], 'amplitudes': []}
        comprehensive_stats[event]['HARPCs'] = {'count': 0, 'velocities': [], 'amplitudes': []}
    
    return comprehensive_stats

def classify_pattern_enhanced(row, sliders, distance_between_sensors):
    """Enhanced pattern classification with new rules and error handling"""
    try:
        # Safely get values with bounds checking
        length_sensors = 0
        if len(row) > 10 and row[10] and row[10].value is not None:
            length_sensors = int(float(row[10].value))
        
        direction = ''
        # Try column 5 (F) for direction - this should contain 'a', 'r', 's'
        if len(row) > 5 and row[5] and row[5].value is not None:
            direction = str(row[5].value).strip()
            # Ensure we only get single character directions
            if direction not in ['a', 'r', 's']:
                direction = ''
            # Ensure we only get single character directions
            if direction in ['a', 'r', 's']:
                direction = direction
            else:
                # Try to extract just the direction character if it's mixed with other data
                for char in direction.lower():
                    if char in ['a', 'r', 's']:
                        direction = char
                        break
                else:
                    direction = ''  # Default if no valid direction found
        
        velocity = 0
        if len(row) > 7 and row[7] and row[7].value is not None:
            try:
                velocity = float(row[7].value)
            except (ValueError, TypeError):
                print(f"DEBUG: Could not convert velocity value '{row[7].value}' to float")
                velocity = 0
        
        # Calculate distance (FIXED: ensure distance_between_sensors is passed properly)
        distance_mm = distance_between_sensors * length_sensors if distance_between_sensors and length_sensors else 0

        print(f"DEBUG: Distance calculation - distance_between_sensors: {distance_between_sensors}, length_sensors: {length_sensors}, distance_mm: {distance_mm}")

        # NEW: Long pattern must be ≥100mm AND ≥5 sensors
        is_long = (distance_mm >= 100) and (length_sensors >= LONG_PATTERN_MINIMUM_SENSORS)
        
        # Get all sensor amplitudes for this pattern (FIXED)
        amplitudes = []
        for col_idx in range(12, min(len(row), 50)):  # Sensor columns start at 12 (after shift)
            if (col_idx < len(row) and row[col_idx] and row[col_idx].value is not None and 
                isinstance(row[col_idx].value, (int, float)) and row[col_idx].value > 0):
                amplitudes.append(float(row[col_idx].value))

        print(f"DEBUG: Pattern amplitudes collected: {len(amplitudes)} sensors, max: {max(amplitudes) if amplitudes else 0}")
        
        # Check for high amplitude (≥3 consecutive sensors ≥100mmHg)
        high_amp_count = count_consecutive_high_amplitude(amplitudes, HIGH_AMPLITUDE_MINIMUM_VALUE)
        is_high_amplitude = (high_amp_count >= HIGH_AMPLITUDE_MINIMUM_PATTERN_LENGTH)

        # FIXED: HAPC/HARPC detection - remove length requirement, focus on amplitude + distance
        is_hapc = is_high_amplitude and (direction == 'a') and (distance_mm >= 100)
        is_harpc = is_high_amplitude and (direction == 'r') and (distance_mm >= 100)

        print(f"DEBUG: HAPC/HARPC check - high_amp_count: {high_amp_count}, direction: {direction}, distance_mm: {distance_mm}, is_hapc: {is_hapc}, is_harpc: {is_harpc}")
        
        pattern_length_category = "Long" if is_long else "Short"
        
        return {
            'length_category': pattern_length_category,
            'direction': direction,
            'velocity': velocity,
            'amplitudes': amplitudes,
            'is_hapc': is_hapc,
            'is_harpc': is_harpc,
            'starting_region': None  # To be filled by region detection
        }
        
    except (ValueError, TypeError, IndexError) as e:
        print(f"Error in classify_pattern_enhanced: {e}")
        # Return default values
        return {
            'length_category': "Short",
            'direction': '',
            'velocity': 0,
            'amplitudes': [],
            'is_hapc': False,
            'is_harpc': False,
            'starting_region': None
        }

def count_consecutive_high_amplitude(amplitudes, threshold):
    """Count consecutive sensors above threshold - FIXED"""
    if not amplitudes or len(amplitudes) < 3:
        return 0
    
    max_consecutive = 0
    current_consecutive = 0
    
    for amp in amplitudes:
        if isinstance(amp, (int, float)) and amp >= threshold:
            current_consecutive += 1
            max_consecutive = max(max_consecutive, current_consecutive)
        else:
            current_consecutive = 0
    
    print(f"DEBUG: HAPC check - amplitudes: {amplitudes[:5]}..., max_consecutive: {max_consecutive}, threshold: {threshold}")
    return max_consecutive

def determine_starting_region(row, sliders):
    """Determine which colon region a pattern starts in based on first active sensor"""
    try:
        # Get slider values
        slider_values = getSliderValues(sliders)
        if not slider_values:
            return "Ascending"  # Default fallback
        
        regions = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
        
        # Remove disabled sections
        active_regions = []
        active_sliders = []
        for i, region in enumerate(regions):
            if region not in disabled_sections and i < len(slider_values):
                active_regions.append(region)
                active_sliders.append(slider_values[i])
        
        if not active_regions:
            return "Ascending"  # Default fallback
        
        # Find the first active sensor (first non-zero sensor value)
        first_active_sensor = None
        for col_idx in range(12, min(len(row), 50)):  # Sensor columns start at index 12
            if (col_idx < len(row) and row[col_idx] and row[col_idx].value is not None and 
                isinstance(row[col_idx].value, (int, float)) and row[col_idx].value > 0):
                # Convert column index to sensor number (col 12 = sensor 1, col 13 = sensor 2, etc.)
                first_active_sensor = col_idx - 11
                break
        
        if first_active_sensor is None:
            return active_regions[0]  # Default to first region
        
        # Find which region contains this sensor (FIXED: proper boundary handling)
        for i, (start_sensor, end_sensor) in enumerate(active_sliders):
            # Fix: Ensure first sensor of region is properly assigned
            if start_sensor <= first_active_sensor <= end_sensor:
                return active_regions[i]

        # If sensor is before first region, assign to first region
        if first_active_sensor < active_sliders[0][0]:
            return active_regions[0]

        # If sensor is after all regions, find the closest one
        for i in range(len(active_sliders) - 1, -1, -1):
            if first_active_sensor >= active_sliders[i][0]:
                return active_regions[i]

        return active_regions[0]  # Default fallback
        
    except (ValueError, TypeError, IndexError) as e:
        print(f"Error in determine_starting_region: {e}")
        return "Ascending"  # Default fallback

def custom_sort(item):
    order = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    return order.index(item)

def remove_disabled_sections(section):
    global disabled_sections
    if section in disabled_sections:
        disabled_sections.remove(section)
    disabled_sections = sorted(disabled_sections, key=custom_sort)

def add_disabled_sections(section):
    global disabled_sections
    if section not in disabled_sections:
        disabled_sections.append(section)
    disabled_sections = sorted(disabled_sections, key=custom_sort)

def reset_disabled_sections():
    global disabled_sections
    disabled_sections = []

def exportToXlsx(data, file_name, sliders, events, settings_sliders, first_event_text):
    #print("DEBUG: Starting exportToXlsx")
    
    try:
        # Split the file path into the base name and extension
        #print("DEBUG: Processing file name")
        base_name, ext = file_name.rsplit('.', 1)
        new_file_name = f"{base_name}_analysis.xlsx"
        #print(f"DEBUG: New file name: {new_file_name}")

        # Write the DataFrame to an Excel file
        #print("DEBUG: Writing DataFrame to Excel")
        data.to_excel(new_file_name, index=False)
        #print("DEBUG: DataFrame written successfully")

        #print("DEBUG: Inserting empty rows")
        insertEmptyRows(new_file_name, 12)
        #print("DEBUG: Empty rows inserted")

        #print("DEBUG: Creating space for comprehensive table")
        create_space_for_comprehensive_table(new_file_name)
        #print("DEBUG: Space created")

        #print("DEBUG: Merging and coloring cells")
        mergeAndColorCells(new_file_name, sliders)
        #print("DEBUG: Cells merged and colored")

        #print("DEBUG: Processing events")
        event_names = []
        for time, event_name in events.items():
            #print(f"DEBUG: Processing event {event_name} at time {time}")
            event_names.append(event_name)
            try:
                total_seconds = time // 10  # Convert deciseconds to seconds
                hour, remainder = divmod(total_seconds, 3600)  # Calculate hours
                minute, second = divmod(remainder, 60)  # Calculate minutes and seconds
                #print(f"DEBUG: Event time converted to {hour}:{minute}:{second}")
                addEventNameAtGivenTime(new_file_name, hour, minute, second, event_name)
                #print(f"DEBUG: Event {event_name} added successfully")
            except Exception as e:
                print(f"DEBUG: Error processing event {event_name}: {e}")
                raise
        
        #print("DEBUG: About to call assignSectionsBasedOnStartSection")
        #print(f"DEBUG: sliders type: {type(sliders)}")
        #print(f"DEBUG: settings_sliders type: {type(settings_sliders)}")
        #print(f"DEBUG: first_event_text: {first_event_text}")
        
        wb = assignSectionsBasedOnStartSection(new_file_name, sliders, event_names, settings_sliders, first_event_text)
        #print("DEBUG: assignSectionsBasedOnStartSection completed")
        
        file_name = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                               filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], 
                                               initialfile=new_file_name)
        wb.save(file_name)
        print(f"Data successfully exported to {new_file_name}")
        
    except Exception as e:
        print(f"DEBUG: Error occurred at: {e}")
        import traceback
        traceback.print_exc()
        print(f"Error exporting data to Excel: {e}")

def getSliderValues(sliders):
    #print(f"DEBUG: getSliderValues called with {len(sliders) if sliders else 0} sliders")
    list_with_slider_tuples = []
    for i in range(len(sliders)):
        #: Processing slider {i}")
        value1 = -1
        value2 = -1
        try:
            slider_values = sliders[i].get()
            #print(f"DEBUG: Slider {i} values: {slider_values}")
            
            if slider_values is None:
                print(f"DEBUG: Slider {i} returned None!")
                continue
                
            for element in slider_values:
                #print(f"DEBUG: Processing element: {element}")
                if value1 == -1:
                    value1 = round(element) if element is not None else 1
                else:
                    value2 = round(element) if element is not None else 1
        except Exception as e:
            print(f"DEBUG: Error processing slider {i}: {e}")
            value1, value2 = 1, 10  # Default values
            
        slider_tuple = (value1, value2)
        #print(f"DEBUG: Slider {i} tuple: {slider_tuple}")
        list_with_slider_tuples.append(slider_tuple)
    #print(f"DEBUG: Final slider values: {list_with_slider_tuples}")
    return list_with_slider_tuples

def mergeAndColorCells(file_name, sliders):
    # Load the workbook and select the active sheet
    wb = load_workbook(file_name)
    ws = wb.active
    
    # Define the sections and their corresponding colors
    sections = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    sliders = getSliderValues(sliders)

    for section in disabled_sections:
        sections.remove(section)
        sliders.pop(0)

    colors = {
        "Ascending": "A9D08E",
        "Transverse": "BDD7EE",
        "Descending": "F8CBAD",
        "Sigmoid": "D9D9D9",
        "Rectum": "B1A0C7"
    }

    # Add empty column between K and L (shift sensor headers right by 1)
    # Insert empty column at L (column 12)
    ws.insert_cols(12)

    # Merge cells and color them for each section (SHIFTED by 1 column)
    for i, (start, end) in enumerate(sliders):
        start_col = get_column_letter(start + 12)  # Changed from +11 to +12
        end_col = get_column_letter(end + 12)      # Changed from +11 to +12
        ws.merge_cells(f'{start_col}71:{end_col}71')
        cell = ws[f'{start_col}71']
        cell.value = sections[i]
        cell.alignment = Alignment(horizontal='center', vertical='center')
        fill = PatternFill(start_color=colors[sections[i]], end_color=colors[sections[i]], fill_type="solid")
        cell.fill = fill

        # Color the individual cells in the merged range
        for col in range(start + 12, end + 13):  # Changed from +11/+12 to +12/+13
            ws[f'{get_column_letter(col)}71'].fill = fill

    # Save the workbook
    wb.save(file_name)

def addEventNameAtGivenTime(file_name, hour, minute, second, event_name):
    # Load the workbook and select the active sheet
    wb = load_workbook(file_name)
    ws = wb.active
    
    #: Looking for insertion point for event {event_name} at {hour}:{minute}:{second}")
    
    # Find the insertion row based on the specified hour, minute, and second
    insertion_row = None
    for row in range(27, ws.max_row + 1):
        try:
            cell_hour = ws.cell(row=row, column=2).value
            cell_minute = ws.cell(row=row, column=3).value
            cell_second = ws.cell(row=row, column=4).value
            
            # Handle None values safely
            if cell_hour is None or cell_minute is None or cell_second is None:
                continue
            
            # Convert to integers safely
            try:
                cell_hour_int = int(cell_hour)
                cell_minute_int = int(cell_minute)
                cell_second_int = int(cell_second)
                hour_int = int(hour)
                minute_int = int(minute)
                second_int = int(second)
            except (ValueError, TypeError):
                continue
            
            # Compare times
            if (cell_hour_int > hour_int or 
                (cell_hour_int == hour_int and cell_minute_int > minute_int) or 
                (cell_hour_int == hour_int and cell_minute_int == minute_int and cell_second_int >= second_int)):
                insertion_row = row
                break
                
        except Exception as e:
            print(f"DEBUG: Error processing row {row}: {e}")
            continue
    
    if insertion_row is None:
        insertion_row = ws.max_row + 1
        #print(f"DEBUG: No suitable insertion point found, inserting at end (row {insertion_row})")
    else:
        print(f"DEBUG: Inserting event at row {insertion_row}")
    
    # Insert a new row at the identified position
    ws.insert_rows(insertion_row)
    
    # Fill in the event details with yellow background
    # Fill in the event details with yellow background (EXTENDED to column K)
    for col in range(1, 12):  # Changed from 11 to 12 to include column K
        cell = ws.cell(row=insertion_row, column=col)
        cell.value = event_name
        cell.fill = PatternFill(start_color=EVENT_COLOR, end_color=EVENT_COLOR, fill_type="solid")
    
    # Fill in the time columns
    ws.cell(row=insertion_row, column=2, value=hour)
    ws.cell(row=insertion_row, column=3, value=minute)
    ws.cell(row=insertion_row, column=4, value=second)

    # Save the workbook
    wb.save(file_name)
    #print(f"DEBUG: Event {event_name} inserted successfully")

def insertEmptyRows(file_name, amount):
    wb = load_workbook(file_name)
    ws = wb.active

    for i in range (amount):
        ws.insert_rows(13 + i)
    
    wb.save(file_name)

def assignSectionsBasedOnStartSection(file_name, sliders, event_names, settings_sliders, first_event_text):
    #print("DEBUG: assignSectionsBasedOnStartSection started")
    #print(f"DEBUG: settings_sliders: {settings_sliders}")
    #print(f"DEBUG: settings_sliders length: {len(settings_sliders) if settings_sliders else 0}")
    
    try:
        if settings_sliders and len(settings_sliders) > 0:
            slider_value = settings_sliders[0].get()
            #print(f"DEBUG: settings_sliders[0].get() returned: {slider_value} (type: {type(slider_value)})")
            distance_between_sensors = int(round(slider_value)) if slider_value is not None else 25
        else:
            #print("DEBUG: No settings_sliders found, using default")
            distance_between_sensors = 25
        #print(f"DEBUG: distance_between_sensors: {distance_between_sensors}")
    except Exception as e:
        print(f"DEBUG: Error getting distance_between_sensors: {e}")
        distance_between_sensors = 25
    
    """Enhanced statistics assignment with comprehensive analysis"""
    wb = load_workbook(file_name)
    ws = wb.active

    # Get event names properly
    # Extract text from CTkEntry widget
    if hasattr(first_event_text, 'get'):
        first_event_value = first_event_text.get().strip() if first_event_text.get() else "Post-Wake"
    else:
        first_event_value = str(first_event_text).strip() if first_event_text else "Post-Wake"

    all_events = [first_event_value]
    all_events.extend(event_names)

    # Initialize comprehensive statistics
    comprehensive_stats = initialize_comprehensive_statistics(all_events)
    
    # KEEP existing counter initialization for backward compatibility
    sections = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    colors = {
        "Ascending": "A9D08E",
        "Transverse": "BDD7EE", 
        "Descending": "F8CBAD",
        "Sigmoid": "D9D9D9",
        "Rectum": "B1A0C7",
        "Ascending tot in Rectum": "81BA5A",
        "Transverse tot in Rectum": "81B2DF",
        "Descending tot in Rectum": "F2A16A",
        "Sigmoid tot in Rectum": "BEBEBE"
    }

    slider_values = getSliderValues(sliders)
    try:
        distance_between_sensors = int(round(settings_sliders[0].get())) if settings_sliders and len(settings_sliders) > 0 else 25
    except (ValueError, TypeError, AttributeError):
        distance_between_sensors = 25  # Default value

    # KEEP existing counter setup
    counters = {}
    length_counters = {}
    high_amplitude_counters = {}
    
    first_event_name = first_event_value
    counters[first_event_name] = {}
    length_counters[first_event_name] = {}
    high_amplitude_counters[first_event_name] = {}
    
    for event_name in event_names:
        counters[event_name] = {}
        length_counters[event_name] = {}  
        high_amplitude_counters[event_name] = {}

    counter_template = {
        "Ascending": 0, "Transverse": 0, "Descending": 0, "Sigmoid": 0, "Rectum": 0,
        "Ascending tot in Rectum": 0, "Transverse tot in Rectum": 0,
        "Descending tot in Rectum": 0, "Sigmoid tot in Rectum": 0,
    }

    length_counter_template = {
        "Long s": 0,    # CHANGED: Removed "aantal"
        "Short s": 0,   # CHANGED: Removed "aantal"  
        "Long r": 0,    # CHANGED: Removed "aantal"
        "Short r": 0,   # CHANGED: Removed "aantal"
        "Long a": 0,    # CHANGED: Removed "aantal"
        "Short a": 0,   # CHANGED: Removed "aantal"
        "cyclic s": 0, "cyclic r": 0, "cyclic a": 0,
    }

    high_amplitude_counters_template = {"HAPCs": 0, "HARPCs": 0}

    # Remove disabled sections
    keys_to_remove = []
    for section in disabled_sections:
        if section in sections:
            sections.remove(section)
            slider_values.pop(0)
        for key in counter_template.keys():
            if section in key:
                keys_to_remove.append(key)

    for key in keys_to_remove:
        if key in counter_template:
            del counter_template[key]

    counter = counter_template.copy()
    length_counter = length_counter_template.copy()
    high_amplitude_counter = high_amplitude_counters_template.copy()
    
    current_event = first_event_name

    # Process each pattern row
    # Process each pattern row
    for row_idx in range(27, ws.max_row + 1):
        try:
            row = [ws.cell(row=row_idx, column=col) for col in range(1, min(ws.max_column + 1, 50))]
            
            # Skip if row is too short or empty
            if len(row) < 12 or not row[0].value:
                continue

            # Check for event markers (skip header rows)
            if (isinstance(row[0].value, str) and 
                not str(row[0].value).isdigit() and 
                row[0].value not in ['Sequence', 'Hour', 'Minute', 'Second', 'Sample']):
                
                # FORCE: Ensure all counters are stored, especially the last event
                for event in all_events:
                    if event not in counters or not counters[event]:
                        print(f"DEBUG: Force storing empty counters for {event}")
                        counters[event] = counter_template.copy()
                        length_counters[event] = length_counter_template.copy() 
                        high_amplitude_counters[event] = high_amplitude_counters_template.copy()

                # Store current counters before switching events (FIXED)
                if current_event:
                    # Apply anti-double-counting to length_counter before storing
                    length_counter["Long a"] = max(0, length_counter["Long a"] - high_amplitude_counter["HAPCs"])
                    length_counter["Long r"] = max(0, length_counter["Long r"] - high_amplitude_counter["HARPCs"])
                    
                    # FIXED: Always store counters
                    counters[current_event] = counter.copy()
                    length_counters[current_event] = length_counter.copy()
                    high_amplitude_counters[current_event] = high_amplitude_counter.copy()
                    
                    print(f"DEBUG: Stored counters for {current_event}:")
                    print(f"  Regional: {counter}")
                    print(f"  Length: {length_counter}")  
                    print(f"  HAPC: {high_amplitude_counter}")
                
                # Switch to new event
                new_event = row[0].value.strip()
                
                # Only switch if this is a valid event we're tracking
                if new_event in all_events:
                    current_event = new_event
                    counter = counter_template.copy()
                    length_counter = length_counter_template.copy()
                    high_amplitude_counter = high_amplitude_counters_template.copy()
                
                continue

            # Skip header rows by checking first cell only
            if (isinstance(row[0].value, str) and 
                row[0].value in ['Sequence', 'Hour', 'Minute', 'Second', 'Sample']):
                continue

            # Skip rows where the length column contains non-numeric data
            if (len(row) > 10 and row[10].value and 
                isinstance(row[10].value, str) and 
                not str(row[10].value).replace('.', '').replace('-', '').isdigit()):
                continue
                
            # Enhanced pattern classification
            classification = classify_pattern_enhanced(row, sliders, distance_between_sensors)
            starting_region = determine_starting_region(row, sliders)
            classification['starting_region'] = starting_region

            print(f"DEBUG Row {row_idx}: Raw direction='{row[6].value if len(row) > 6 else 'N/A'}', "
                f"Processed direction='{classification['direction']}', "
                f"Length='{classification['length_category']}', "
                f"Velocity='{classification['velocity']}', "
                f"Region='{starting_region}', Event='{current_event}'")

            # Fill Start Region column (Column K)
            region_cell = ws.cell(row=row_idx, column=11, value=starting_region)
            
            # Color the Start Region cell based on region
            region_colors = {
                "Ascending": "A9D08E",
                "Transverse": "BDD7EE", 
                "Descending": "F8CBAD",
                "Sigmoid": "D9D9D9",
                "Rectum": "B1A0C7"
            }
            if starting_region in region_colors:
                region_cell.fill = PatternFill(start_color=region_colors[starting_region], 
                                            end_color=region_colors[starting_region], fill_type="solid")
            
            # Update comprehensive statistics
            update_comprehensive_stats(comprehensive_stats, classification, current_event, row, sliders)
            
            # Handle HAPCs/HARPCs for BOTH old and new counters (FIXED)
            if classification['is_hapc']:
                high_amplitude_counter["HAPCs"] += 1
                comprehensive_stats[current_event]['HAPCs']['count'] += 1
                if classification['velocity'] != 0:
                    comprehensive_stats[current_event]['HAPCs']['velocities'].append(classification['velocity'])
                if classification['amplitudes']:
                    comprehensive_stats[current_event]['HAPCs']['amplitudes'].extend(classification['amplitudes'])
                
                print(f"DEBUG: HAPC detected! Count now: {high_amplitude_counter['HAPCs']}")
                
                # Color HAPC sensor cells GREEN (#92D050) - FIXED for column shift
                for col_idx in range(13, len(row) + 13):  # Adjust for empty column insertion
                    if (col_idx - 13 < len(row) and row[col_idx - 13].value and 
                        isinstance(row[col_idx - 13].value, (int, float)) and 
                        row[col_idx - 13].value >= HIGH_AMPLITUDE_MINIMUM_VALUE):
                        sensor_cell = ws.cell(row=row_idx, column=col_idx)
                        sensor_cell.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
                        print(f"DEBUG: Colored HAPC cell at row {row_idx}, col {col_idx}")
                            
            elif classification['is_harpc']:
                high_amplitude_counter["HARPCs"] += 1
                comprehensive_stats[current_event]['HARPCs']['count'] += 1
                if classification['velocity'] != 0:
                    comprehensive_stats[current_event]['HARPCs']['velocities'].append(classification['velocity'])
                if classification['amplitudes']:
                    comprehensive_stats[current_event]['HARPCs']['amplitudes'].extend(classification['amplitudes'])
                
                print(f"DEBUG: HARPC detected! Count now: {high_amplitude_counter['HARPCs']}")
                
                # Color HARPC sensor cells RED (#FF0000) - FIXED for column shift
                for col_idx in range(13, len(row) + 13):  # Adjust for empty column insertion
                    if (col_idx - 13 < len(row) and row[col_idx - 13].value and 
                        isinstance(row[col_idx - 13].value, (int, float)) and 
                        row[col_idx - 13].value >= HIGH_AMPLITUDE_MINIMUM_VALUE):
                        sensor_cell = ws.cell(row=row_idx, column=col_idx)
                        sensor_cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        print(f"DEBUG: Colored HARPC cell at row {row_idx}, col {col_idx}")

            # FIXED: Use classification results for pattern counting
            pattern = classification['direction']
            length = 0
            if len(row) > 10 and row[10].value is not None:
                try:
                    length = int(float(str(row[10].value)))
                except (ValueError, TypeError):
                    continue
            
            # FIXED: Update length counters using classification results
            if pattern and classification['length_category']:
                pattern_type = classification['length_category']  # "Long" or "Short"
                counter_key = f"{pattern_type} {pattern}"
                
                if counter_key in length_counter:
                    length_counter[counter_key] += 1
                    print(f"DEBUG: Updated length counter {counter_key} to {length_counter[counter_key]} for event {current_event}")
                else:
                    print(f"DEBUG: Length counter key '{counter_key}' not found. Available keys: {list(length_counter.keys())}")

            # Regional counting logic for old table (COMPLETELY FIXED)
            if starting_region and starting_region in counter:
                counter[starting_region] += 1
                print(f"DEBUG: Regional counter - Added to {starting_region}, now {counter[starting_region]}")
                
                # Color the length cell using region colors
                length_cell = ws.cell(row=row_idx, column=10)  # Column J (Length)
                fill = PatternFill(start_color=colors[starting_region], end_color=colors[starting_region], fill_type="solid")
                length_cell.fill = fill
                
                # Check for pan-colonic pattern using comprehensive stats detection
                if is_pan_colonic_pattern(row, sliders, starting_region):
                    pan_colonic_key = f'{starting_region} tot in Rectum'
                    if pan_colonic_key in counter:
                        counter[pan_colonic_key] += 1
                        print(f"DEBUG: Pan-colonic counter - Added to {pan_colonic_key}, now {counter[pan_colonic_key]}")
        except Exception as e:
            print(f"DEBUG: Error processing row {row_idx}: {e}")
            continue

    # Handle final event
    if counters[current_event] == {}:
        # Subtract pan-colonic from regional to avoid double-counting
        for section in sections[:-1] if len(sections) > 1 else []:
            section_key = f'{section} tot in Rectum'
            if section_key in counter:
                counter[section] = max(0, counter[section] - counter[section_key])

        # Apply anti-double-counting
        length_counter["Long a"] = max(0, length_counter["Long a"] - high_amplitude_counter["HAPCs"])
        length_counter["Long r"] = max(0, length_counter["Long r"] - high_amplitude_counter["HARPCs"])

        counters[current_event] = counter.copy()
        length_counters[current_event] = length_counter.copy()
        high_amplitude_counters[current_event] = high_amplitude_counter.copy()

    # ENSURE final event counters are stored
    if current_event and (current_event not in counters or all(v == 0 for v in counters[current_event].values())):
        # Apply final anti-double-counting
        length_counter["Long a"] = max(0, length_counter["Long a"] - high_amplitude_counter["HAPCs"])
        length_counter["Long r"] = max(0, length_counter["Long r"] - high_amplitude_counter["HARPCs"])
        
        counters[current_event] = counter.copy()
        length_counters[current_event] = length_counter.copy()
        high_amplitude_counters[current_event] = high_amplitude_counter.copy()
        
        print(f"DEBUG: Final storage for {current_event}:")
        print(f"  Regional: {counter}")
        print(f"  Length: {length_counter}")
        print(f"  HAPC: {high_amplitude_counter}")

    # DEBUG: Print final counter values before writing to Excel
    print("DEBUG: Final counter values:")
    for event, event_counters in counters.items():
        print(f"  Event {event}: {event_counters}")
    for event, length_counters_data in length_counters.items():
        print(f"  Length counters {event}: {length_counters_data}")
    for event, hapc_counters in high_amplitude_counters.items():
        print(f"  HAPC counters {event}: {hapc_counters}")

    # Apply final anti-double-counting corrections to comprehensive stats
    apply_hapc_harpc_corrections(comprehensive_stats)
    
    # NEW: Create the comprehensive analysis table
    create_comprehensive_analysis_table(wb, comprehensive_stats, all_events)
    
    # KEEP existing summary table creation code
    # [Rest of existing function for creating the summary in columns S-W]

    # RESTORE existing summary table creation code
    #Section names added to the left of table
    row = 3
    for section in counter_template.keys():
        ws.cell(row=row, column=19, value=section)
        fill = PatternFill(start_color=colors[section], end_color=colors[section], fill_type="solid")
        ws.cell(row=row, column=19).fill = fill
        row += 1
    row += 1
    
    #Pattern types (eg. Long s, Short s) added to the left of table
    length_counter_row_start = row
    for pattern in list(chain(length_counter_template.keys(), high_amplitude_counters_template.keys())):
        ws.cell(row=row, column=19, value=pattern)
        
        # Use specific colors for HAPCs and HARPCs
        if pattern == "HAPCs":
            fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
        elif pattern == "HARPCs":
            fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        else:
            fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
        
        ws.cell(row=row, column=19).fill = fill
        row += 1

    # DEBUG: Print what we're about to write to Excel
    print("DEBUG: About to write to Excel:")
    print(f"  counters keys: {list(counters.keys())}")
    print(f"  length_counters keys: {list(length_counters.keys())}")
    print(f"  high_amplitude_counters keys: {list(high_amplitude_counters.keys())}")

    for event in counters:
        print(f"  {event} regional: {counters[event]}")
        print(f"  {event} length: {length_counters[event]}")
        print(f"  {event} hapc: {high_amplitude_counters[event]}")
    
    #Fill in the counters
    column = 20
    for (event,value) in counters.items():
        row = 2
        ws.cell(row=row, column=column, value=event)
        fill = PatternFill(start_color=EVENT_COLOR, end_color=EVENT_COLOR, fill_type="solid")
        ws.cell(row=row, column=column).fill = fill
        row += 1
        for (section,value) in value.items():
            ws.cell(row=row, column=column, value=value)
            fill = PatternFill(start_color=colors[section], end_color=colors[section], fill_type="solid")
            ws.cell(row=row, column=column).fill = fill
            row += 1
        column += 1
    
    #Fill in the length counters
    column = 20
    for (event,value) in length_counters.items():
        row = length_counter_row_start
        for pattern in value.values():
            ws.cell(row=row, column=column, value=pattern)
            row += 1
        column += 1
    
    column = 20
    high_amplitude_start = row
    for (event,value) in high_amplitude_counters.items():
        row = high_amplitude_start
        for pattern in value.values():
            ws.cell(row=row, column=column, value=pattern)
            row += 1
        column += 1
    
    return wb

def update_comprehensive_stats(comprehensive_stats, classification, current_event, row, sliders):
    """Update comprehensive statistics with pattern data"""
    
    # Skip if current_event is not in our comprehensive_stats
    if current_event not in comprehensive_stats:
        print(f"DEBUG: Skipping stats update for unknown event: {current_event}")
        return
    
    pattern_key = f"{classification['length_category']} {classification['direction']}"
    starting_region = classification['starting_region']
    
    print(f"DEBUG: Updating stats for event={current_event}, pattern={pattern_key}, region={starting_region}")
    
    # Skip if pattern type not in our tracking
    if pattern_key not in comprehensive_stats[current_event]:
        print(f"DEBUG: Pattern key {pattern_key} not found, available keys: {list(comprehensive_stats[current_event].keys())}")
        return
    
    # Update regional stats
    if starting_region in comprehensive_stats[current_event][pattern_key]:
        stats = comprehensive_stats[current_event][pattern_key][starting_region]
        stats['count'] += 1
        if classification['velocity'] != 0:
            stats['velocities'].append(classification['velocity'])
        if classification['amplitudes']:
            stats['amplitudes'].extend(classification['amplitudes'])
        print(f"DEBUG: Updated {starting_region} count to {stats['count']}")
    else:
        print(f"DEBUG: Region {starting_region} not found in pattern {pattern_key}")
    
    # Check for pan-colonic patterns (region-range stats)
    if is_pan_colonic_pattern(row, sliders, starting_region):
        region_range = f"{starting_region} - Rectum"
        if region_range in comprehensive_stats[current_event][pattern_key]:
            stats = comprehensive_stats[current_event][pattern_key][region_range]
            stats['count'] += 1
            if classification['velocity'] != 0:
                stats['velocities'].append(classification['velocity'])
            if classification['amplitudes']:
                stats['amplitudes'].extend(classification['amplitudes'])
            print(f"DEBUG: Updated pan-colonic {region_range} count to {stats['count']}")
    
    # Update totals
    if 'Total' in comprehensive_stats[current_event][pattern_key]:
        stats = comprehensive_stats[current_event][pattern_key]['Total']
        stats['count'] += 1
        if classification['velocity'] != 0:
            stats['velocities'].append(classification['velocity'])
        if classification['amplitudes']:
            stats['amplitudes'].extend(classification['amplitudes'])
        print(f"DEBUG: Updated Total {pattern_key} count to {stats['count']}")

def apply_hapc_harpc_corrections(comprehensive_stats):
    """Apply anti-double-counting for HAPCs/HARPCs"""
    for event in comprehensive_stats:
        hapc_count = comprehensive_stats[event]['HAPCs']['count']
        harpc_count = comprehensive_stats[event]['HARPCs']['count']
        
        # Subtract HAPCs from Long a counts
        if hapc_count > 0 and 'Long a' in comprehensive_stats[event]:
            for region_key in comprehensive_stats[event]['Long a']:
                if comprehensive_stats[event]['Long a'][region_key]['count'] > 0:
                    # Distribute the subtraction proportionally or from totals
                    comprehensive_stats[event]['Long a']['Total']['count'] -= hapc_count
                    break
        
        # Subtract HARPCs from Long r counts  
        if harpc_count > 0 and 'Long r' in comprehensive_stats[event]:
            for region_key in comprehensive_stats[event]['Long r']:
                if comprehensive_stats[event]['Long r'][region_key]['count'] > 0:
                    comprehensive_stats[event]['Long r']['Total']['count'] -= harpc_count
                    break

def is_pan_colonic_pattern(row, sliders, starting_region):
    """Check if pattern propagates to rectum - FIXED"""
    slider_values = getSliderValues(sliders)
    if len(slider_values) < 5:  # No rectum configured
        return False
    
    rectum_start, rectum_end = slider_values[4]  # Get rectum sensor range
    
    # Check if there are active sensors in rectum region
    has_rectum_activity = False
    for col in range(12, min(len(row), rectum_end + 12)):  # Account for column shift
        sensor_num = col - 11  # Convert to sensor number
        if (rectum_start <= sensor_num <= rectum_end and 
            col < len(row) and row[col].value is not None and 
            isinstance(row[col].value, (int, float)) and row[col].value > 0):
            has_rectum_activity = True
            break
    
    # Also check if there's activity in the starting region
    starting_region_idx = None
    regions = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    if starting_region in regions:
        starting_region_idx = regions.index(starting_region)
    
    if starting_region_idx is not None and starting_region_idx < len(slider_values):
        start_sensors, end_sensors = slider_values[starting_region_idx]
        has_starting_activity = any(
            (start_sensors <= (col - 11) <= end_sensors and 
             col < len(row) and row[col].value is not None and 
             isinstance(row[col].value, (int, float)) and row[col].value > 0)
            for col in range(12, min(len(row), end_sensors + 12))
        )
        
        result = has_rectum_activity and has_starting_activity
        if result:
            print(f"DEBUG: Pan-colonic pattern detected: {starting_region} to Rectum")
        return result
    
    return False

def create_space_for_comprehensive_table(file_name):
    """Move existing sequence/pan-colonic tables down to make space for comprehensive table"""
    wb = load_workbook(file_name)
    ws = wb.active
    
    # The comprehensive table needs rows 2-68 in columns AA-BC
    # Current sequence table starts around row 25, needs to move to row 71
    
    # Insert 46 rows at position 25 to push everything down
    # This moves the sequence table from row 25 to row 71 (25 + 46 = 71)
    rows_to_insert = 71 - 25  # = 46 rows
    
    ws.insert_rows(25, rows_to_insert)
    
    wb.save(file_name)

def create_comprehensive_analysis_table(wb, comprehensive_stats, event_names):
    """Create the comprehensive analysis table at AA2:BC68"""
    ws = wb.active
    
    # Table position
    start_row = 2
    start_col = 27  # Column AA (0-indexed: 26)

    print(f"DEBUG: Creating comprehensive table at row {start_row}, col {start_col}")

    # Check for existing merged cells that actually conflict with our data area (rows 2-68)
    ranges_to_unmerge = []
    for merged_range in ws.merged_cells.ranges:
        # Only unmerge if it conflicts with our data area (rows 2-68), not headers (row 71)
        if (merged_range.min_row >= start_row and merged_range.max_row <= 68 and
            merged_range.min_col >= start_col and merged_range.max_col <= start_col + 30):
            ranges_to_unmerge.append(merged_range)

    # Unmerge conflicting cells (should be none now)
    for range_to_unmerge in ranges_to_unmerge:
        print(f"DEBUG: Unmerging cells in range {range_to_unmerge}")
        ws.unmerge_cells(str(range_to_unmerge))
    
    # Create headers
    create_comprehensive_table_headers(ws, start_row, start_col, event_names)
    
    # Create pattern type rows
    current_row = start_row + 2  # Start at row 4 (after headers)
    
    pattern_types = ["Long s", "Short s", "Long r", "Short r", "Long a", "Short a"]
    regions = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    region_ranges = ["Ascending - Rectum", "Transverse - Rectum", 
                    "Descending - Rectum", "Sigmoid-Rectum"]
    
    # Create all pattern type rows
    for pattern_type in pattern_types:
        # Regional breakdown
        for region in regions:
            row_label = f"{pattern_type} {region}"
            create_pattern_row(ws, current_row, start_col, row_label, 
                             comprehensive_stats, event_names, pattern_type, region)
            current_row += 1
        
        # Region-range breakdown  
        for region_range in region_ranges:
            row_label = f"{pattern_type} {region_range}"
            create_pattern_row(ws, current_row, start_col, row_label, 
                             comprehensive_stats, event_names, pattern_type, region_range)
            current_row += 1
        
        # Total row
        row_label = f"Total {pattern_type}"
        create_pattern_row(ws, current_row, start_col, row_label, 
                         comprehensive_stats, event_names, pattern_type, 'Total')
        current_row += 1
    
    # Cyclic patterns
    for cyclic_type in ['cyclic s', 'cyclic r', 'cyclic a']:
        create_pattern_row(ws, current_row, start_col, cyclic_type, 
                         comprehensive_stats, event_names, cyclic_type, None)
        current_row += 1
    
    # HAPCs and HARPCs
    create_pattern_row(ws, current_row, start_col, "HAPCs", 
                     comprehensive_stats, event_names, 'HAPCs', None)
    current_row += 1
    create_pattern_row(ws, current_row, start_col, "HARPCs", 
                     comprehensive_stats, event_names, 'HARPCs', None)
    
    # Apply formatting
    apply_comprehensive_table_formatting(ws, start_row, start_col, event_names)

def create_comprehensive_table_headers(ws, start_row, start_col, event_names):
    """Create headers for the comprehensive table"""
    
    # Row 1: Event names (Post-Wake, Start FODMAP, etc.)
    col_offset = 1  # Skip the label column
    for i, event in enumerate(event_names):
        event_start_col = start_col + col_offset
        
        # Merge 7 columns for each event (count, min/max/mean vel, min/max/mean amp)
        ws.merge_cells(start_row=start_row, start_column=event_start_col, 
                      end_row=start_row, end_column=event_start_col + 6)
        
        # Set event name
        ws.cell(row=start_row, column=event_start_col, value=event)
        
        col_offset += 7
    
    # Row 2: Metric headers
    ws.cell(row=start_row + 1, column=start_col, value="Type of event")
    
    col_offset = 1
    metric_headers = ["Number of patterns", "min vel", "max vel", "mean vel", 
                     "min amp", "max amp", "mean amp"]
    
    for event in event_names:
        for j, metric in enumerate(metric_headers):
            ws.cell(row=start_row + 1, column=start_col + col_offset + j, value=metric)
        col_offset += 7

def create_pattern_row(ws, row, start_col, row_label, comprehensive_stats, 
                      event_names, pattern_type, region_key):
    """Create a single pattern row with all metrics"""
    
    try:
        # Check if the cell is merged and skip if it is
        cell_ref = f"{get_column_letter(start_col)}{row}"
        if any(cell_ref in merged_range for merged_range in ws.merged_cells.ranges):
            print(f"DEBUG: Skipping merged cell at {cell_ref}")
            return
        
        # Set row label
        ws.cell(row=row, column=start_col, value=row_label)
        
        col_offset = 1
        for event in event_names:
            # Get statistics for this pattern type, event, and region
            stats = get_pattern_statistics(comprehensive_stats, event, pattern_type, region_key)
            
            # Column 1: Count
            ws.cell(row=row, column=start_col + col_offset, value=stats['count'])
            
            # Columns 2-4: Velocity metrics (min, max, mean)
            if stats['velocities']:
                ws.cell(row=row, column=start_col + col_offset + 1, 
                       value=round(min(stats['velocities']), 2))
                ws.cell(row=row, column=start_col + col_offset + 2, 
                       value=round(max(stats['velocities']), 2))
                ws.cell(row=row, column=start_col + col_offset + 3, 
                       value=round(sum(stats['velocities']) / len(stats['velocities']), 2))
            # else: leave blank for empty cells
            
            # Columns 5-7: Amplitude metrics (min, max, mean)  
            if stats['amplitudes']:
                ws.cell(row=row, column=start_col + col_offset + 4, 
                       value=round(min(stats['amplitudes']), 1))
                ws.cell(row=row, column=start_col + col_offset + 5, 
                       value=round(max(stats['amplitudes']), 1))
                ws.cell(row=row, column=start_col + col_offset + 6, 
                       value=round(sum(stats['amplitudes']) / len(stats['amplitudes']), 1))
            # else: leave blank for empty cells
            
            col_offset += 7
            
    except Exception as e:
        print(f"DEBUG: Error creating pattern row {row}: {e}")
        # Continue without crashing

def get_pattern_statistics(comprehensive_stats, event, pattern_type, region_key):
    """Get statistics for a specific pattern type, event, and region"""
    
    default_stats = {'count': 0, 'velocities': [], 'amplitudes': []}
    
    if event not in comprehensive_stats:
        print(f"DEBUG: Event {event} not found in comprehensive_stats")
        return default_stats
    
    # Handle special cases (HAPCs, HARPCs, cyclic)
    if pattern_type in ['HAPCs', 'HARPCs', 'cyclic s', 'cyclic r', 'cyclic a']:
        if pattern_type in comprehensive_stats[event]:
            result = comprehensive_stats[event][pattern_type]
            print(f"DEBUG: Special pattern {pattern_type} for {event}: count={result.get('count', 0)}")
            return result
        else:
            print(f"DEBUG: Special pattern {pattern_type} not found for {event}")
            return default_stats
    
    # Handle regular pattern types
    if pattern_type not in comprehensive_stats[event]:
        print(f"DEBUG: Pattern {pattern_type} not found in event {event}")
        return default_stats
    
    if region_key is None or region_key not in comprehensive_stats[event][pattern_type]:
        print(f"DEBUG: Region {region_key} not found in pattern {pattern_type} for event {event}")
        return default_stats
    
    result = comprehensive_stats[event][pattern_type][region_key]
    print(f"DEBUG: Found {pattern_type} {region_key} for {event}: count={result.get('count', 0)}")
    return result

def apply_comprehensive_table_formatting(ws, start_row, start_col, event_names):
    """Apply colors and formatting to the comprehensive table"""
    
    # Color scheme
    colors = {
        "Ascending": "A9D08E",
        "Transverse": "BDD7EE", 
        "Descending": "F8CBAD",
        "Sigmoid": "D9D9D9",
        "Rectum": "B1A0C7",
        "Ascending - Rectum": "81BA5A", 
        "Transverse - Rectum": "81B2DF",
        "Descending - Rectum": "F2A16A",
        "Sigmoid-Rectum": "BEBEBE"
    }
    
    # Header row colors (Row 2: Events)
    col_offset = 1
    for i, event in enumerate(event_names):
        # Alternate between two event colors
        event_color = ALTERNATING_EVENT_COLORS[i % 2]
        event_start_col = start_col + col_offset
        
        # Color the merged event header
        for col in range(event_start_col, event_start_col + 7):
            cell = ws.cell(row=start_row, column=col)
            cell.fill = PatternFill(start_color=event_color, end_color=event_color, fill_type="solid")
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        col_offset += 7
    
    # Metric header row colors (Row 3: Light gray)
    metric_gray = "D9D9D9"
    for col in range(start_col, start_col + 1 + len(event_names) * 7):
        cell = ws.cell(row=start_row + 1, column=col)
        cell.fill = PatternFill(start_color=metric_gray, end_color=metric_gray, fill_type="solid")
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Pattern type row colors (Column AA: Based on region)
    current_row = start_row + 2
    pattern_types = ["Long s", "Short s", "Long r", "Short r", "Long a", "Short a"]
    regions = ["Ascending", "Transverse", "Descending", "Sigmoid", "Rectum"]
    region_ranges = ["Ascending - Rectum", "Transverse - Rectum", 
                    "Descending - Rectum", "Sigmoid-Rectum"]
    
    for pattern_type in pattern_types:
        # Color regional rows
        for region in regions:
            if region in colors:
                cell = ws.cell(row=current_row, column=start_col)
                cell.fill = PatternFill(start_color=colors[region], end_color=colors[region], fill_type="solid")
            current_row += 1
        
        # Color region-range rows
        for region_range in region_ranges:
            if region_range in colors:
                cell = ws.cell(row=current_row, column=start_col)
                cell.fill = PatternFill(start_color=colors[region_range], end_color=colors[region_range], fill_type="solid")
            current_row += 1
        
        # Total rows - leave white (no color)
        current_row += 1
    
    # Cyclic and HAPC/HARPC rows - leave white
    current_row += 5  # Skip cyclic s, r, a, HAPCs, HARPCs

    # Color special pattern types
    special_colors = {
        "cyclic s": "F4B084",
        "cyclic r": "F4B084", 
        "cyclic a": "F4B084",
        "HAPCs": "92D050",
        "HARPCs": "FF0000"
    }
    
    # Apply colors to specific rows (cyclic patterns start at row 64, HAPCs at 67, HARPCs at 68)
    special_pattern_rows = {
        64: ("cyclic s", "F4B084"),   # AA64
        65: ("cyclic r", "F4B084"),   # AA65  
        66: ("cyclic a", "F4B084"),   # AA66
        67: ("HAPCs", "92D050"),      # AA67
        68: ("HARPCs", "FF0000")      # AA68
    }

    for row_num, (pattern_type, color) in special_pattern_rows.items():
        cell = ws.cell(row=row_num, column=start_col)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")