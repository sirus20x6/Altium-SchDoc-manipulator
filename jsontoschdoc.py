import json
import olefile
import struct
import logging

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def handle_header(stream, record):
    """
    Handles the Header record, which sets up the initial metadata for the schematic file.
    This includes information about the version, unique identifiers, and overall structure weight.
    """
    header = record.get('HEADER', "Protel for Windows - Schematic Capture Binary File Version 5.0").encode('utf-8')
    weight = int(record.get('WEIGHT', 0))  # Default to 0 if not specified
    minor_version = int(record.get('MINORVERSION', 0))  # Default version
    unique_id = record.get('UNIQUEID', '').encode('utf-8')  # Optional unique identifier
    
    data = f"|HEADER={header}|WEIGHT={weight}|MINORVERSION={minor_version}|UNIQUEID={unique_id}".encode('utf-8')
    stream.write(data + b"\n")

def handle_component(stream, record):
    """
    Writes a component record to the stream based on its properties.
    """
    # Start the record with its type
    data = b"|RECORD=1"

    # Mandatory fields
    libref = record.get('LIBREFERENCE', '').encode('ascii')
    data += f"|LIBREFERENCE={libref}".encode('ascii')

    # Optional fields with UTF-8 support
    if 'COMPONENTDESCRIPTION' in record:
        description = record['COMPONENTDESCRIPTION'].encode('ascii')
        utf8_description = record.get('%UTF8%COMPONENTDESCRIPTION', description).encode('utf-8')
        data += f"|COMPONENTDESCRIPTION={description}|%UTF8%COMPONENTDESCRIPTION={utf8_description}".encode('ascii')

    # Integer fields
    part_count = int(record.get('PARTCOUNT', 1))
    display_mode_count = int(record.get('DISPLAYMODECOUNT', 0))
    index_in_sheet = int(record.get('INDEXINSHEET', -1))
    current_part_id = int(record.get('CURRENTPARTID', -1))
    data += f"|PARTCOUNT={part_count}|DISPLAYMODECOUNT={display_mode_count}|INDEXINSHEET={index_in_sheet}|CURRENTPARTID={current_part_id}".encode('ascii')

    # Location fields
    location_x = int(record.get('LOCATION.X', 0))
    location_y = int(record.get('LOCATION.Y', 0))
    data += f"|LOCATION.X={location_x}|LOCATION.Y={location_y}".encode('ascii')

    # Optional boolean fields
    is_mirrored = 'T' if record.get('ISMIRRORED', False) else 'F'
    partid_locked = 'T' if record.get('PARTIDLOCKED', False) else 'F'
    designator_locked = 'T' if record.get('DESIGNATORLOCKED', False) else 'F'
    pins_moveable = 'T' if record.get('PINSMOVEABLE', False) else 'F'
    data += f"|ISMIRRORED={is_mirrored}|PARTIDLOCKED={partid_locked}|DESIGNATORLOCKED={designator_locked}|PINSMOVEABLE={pins_moveable}".encode('ascii')

    # Color fields
    area_color = int(record.get('AREACOLOR', 11599871))
    color = int(record.get('COLOR', 128))
    data += f"|AREACOLOR={area_color}|COLOR={color}".encode('ascii')

    # Optional paths and IDs
    for key in ['LIBRARYPATH', 'SOURCELIBRARYNAME', 'SHEETPARTFILENAME', 'TARGETFILENAME', 'UNIQUEID']:
        if key in record:
            value = record[key].encode('ascii')
            data += f"|{key}={value}".encode('ascii')

    # Additional optional fields
    for key in ['NOTUSEDBTABLENAME', 'DESIGNITEMID', 'DATABASETABLENAME', 'ALIASLIST']:
        if key in record:
            value = record[key].encode('ascii')
            data += f"|{key}={value}".encode('ascii')

    # Write the compiled data to the stream
    stream.write(data + b"\n")  # Assuming newline as a record separator

def handle_pin(stream, record):
    """
    Writes a pin record to the stream based on its properties.
    """
    # Start the record with its type
    data = b"|RECORD=2"

    # Mandatory fields
    owner_index = int(record.get('OWNERINDEX', 0))
    owner_part_id = int(record.get('OWNERPARTID', -1))
    pin_length = int(record.get('PINLENGTH', 0))
    location_x = int(record.get('LOCATION.X', 0))
    location_y = int(record.get('LOCATION.Y', 0))
    electrical = int(record.get('ELECTRICAL', 0))
    data += f"|OWNERINDEX={owner_index}|OWNERPARTID={owner_part_id}|PINLENGTH={pin_length}|LOCATION.X={location_x}|LOCATION.Y={location_y}|ELECTRICAL={electrical}".encode('ascii')

    # Optional fields
    if 'OWNERPARTDISPLAYMODE' in record:
        display_mode = int(record['OWNERPARTDISPLAYMODE'])
        data += f"|OWNERPARTDISPLAYMODE={display_mode}".encode('ascii')

    if 'DESCRIPTION' in record:
        description = record['DESCRIPTION'].encode('ascii')
        data += f"|DESCRIPTION={description}".encode('ascii')

    # Symbol properties
    symbol_outer = int(record.get('SYMBOL_OUTER', 0))
    symbol_outer_edge = int(record.get('SYMBOL_OUTEREDGE', 0))
    symbol_inner_edge = int(record.get('SYMBOL_INNEREDGE', 0))
    data += f"|SYMBOL_OUTER={symbol_outer}|SYMBOL_OUTEREDGE={symbol_outer_edge}|SYMBOL_INNEREDGE={symbol_inner_edge}".encode('ascii')

    # Name and designator handling
    if 'NAME' in record:
        name = record['NAME'].encode('ascii')
        name_custom_font_id = int(record.get('NAME_CUSTOMFONTID', 4))
        name_custom_position_margin = int(record.get('NAME_CUSTOMPOSITION_MARGIN', -7))
        pinname_position_conglomerate = int(record.get('PINNAME_POSITIONCONGLOMERATE', 0))
        data += f"|NAME={name}|NAME_CUSTOMFONTID={name_custom_font_id}|NAME_CUSTOMPOSITION_MARGIN={name_custom_position_margin}|PINNAME_POSITIONCONGLOMERATE={pinname_position_conglomerate}".encode('ascii')

    if 'DESIGNATOR' in record:
        designator = record['DESIGNATOR'].encode('ascii')
        designator_custom_position_margin = int(record.get('DESIGNATOR_CUSTOMPOSITION_MARGIN', 9))
        pindesignator_position_conglomerate = int(record.get('PINDESIGNATOR_POSITIONCONGLOMERATE', 0))
        data += f"|DESIGNATOR={designator}|DESIGNATOR_CUSTOMPOSITION_MARGIN={designator_custom_position_margin}|PINDESIGNATOR_POSITIONCONGLOMERATE={pindesignator_position_conglomerate}".encode('ascii')

    # Swap ID and UTF-8 handling for optional parts
    if 'SWAPIDPIN' in record:
        swapidpin = record['SWAPIDPIN'].encode('ascii')
        data += f"|SWAPIDPIN={swapidpin}".encode('ascii')

    if '%UTF8%SWAPIDPART' in record:
        utf8_swapidpart = record['%UTF8%SWAPIDPART'].encode('utf-8')
        data += f"|%UTF8%SWAPIDPART={utf8_swapidpart}".encode('ascii')

    # Write the compiled data to the stream
    stream.write(data + b"\n")  # Assuming newline as a record separator

def handle_pin(stream, record):
    """
    Writes a pin record to the stream based on its properties.
    """
    # Start the record with its type
    data = b"|RECORD=2"

    # Mandatory fields
    owner_index = int(record.get('OWNERINDEX', 0))
    owner_part_id = int(record.get('OWNERPARTID', -1))
    pin_length = int(record.get('PINLENGTH', 0))
    location_x = int(record.get('LOCATION.X', 0))
    location_y = int(record.get('LOCATION.Y', 0))
    electrical = int(record.get('ELECTRICAL', 0))
    data += f"|OWNERINDEX={owner_index}|OWNERPARTID={owner_part_id}|PINLENGTH={pin_length}|LOCATION.X={location_x}|LOCATION.Y={location_y}|ELECTRICAL={electrical}".encode('ascii')

    # Optional fields
    if 'OWNERPARTDISPLAYMODE' in record:
        display_mode = int(record['OWNERPARTDISPLAYMODE'])
        data += f"|OWNERPARTDISPLAYMODE={display_mode}".encode('ascii')

    if 'DESCRIPTION' in record:
        description = record['DESCRIPTION'].encode('ascii')
        data += f"|DESCRIPTION={description}".encode('ascii')

    # Symbol properties
    symbol_outer = int(record.get('SYMBOL_OUTER', 0))
    symbol_outer_edge = int(record.get('SYMBOL_OUTEREDGE', 0))
    symbol_inner_edge = int(record.get('SYMBOL_INNEREDGE', 0))
    data += f"|SYMBOL_OUTER={symbol_outer}|SYMBOL_OUTEREDGE={symbol_outer_edge}|SYMBOL_INNEREDGE={symbol_inner_edge}".encode('ascii')

    # Name and designator handling
    if 'NAME' in record:
        name = record['NAME'].encode('ascii')
        name_custom_font_id = int(record.get('NAME_CUSTOMFONTID', 4))
        name_custom_position_margin = int(record.get('NAME_CUSTOMPOSITION_MARGIN', -7))
        pinname_position_conglomerate = int(record.get('PINNAME_POSITIONCONGLOMERATE', 0))
        data += f"|NAME={name}|NAME_CUSTOMFONTID={name_custom_font_id}|NAME_CUSTOMPOSITION_MARGIN={name_custom_position_margin}|PINNAME_POSITIONCONGLOMERATE={pinname_position_conglomerate}".encode('ascii')

    if 'DESIGNATOR' in record:
        designator = record['DESIGNATOR'].encode('ascii')
        designator_custom_position_margin = int(record.get('DESIGNATOR_CUSTOMPOSITION_MARGIN', 9))
        pindesignator_position_conglomerate = int(record.get('PINDESIGNATOR_POSITIONCONGLOMERATE', 0))
        data += f"|DESIGNATOR={designator}|DESIGNATOR_CUSTOMPOSITION_MARGIN={designator_custom_position_margin}|PINDESIGNATOR_POSITIONCONGLOMERATE={pindesignator_position_conglomerate}".encode('ascii')

    # Swap ID and UTF-8 handling for optional parts
    if 'SWAPIDPIN' in record:
        swapidpin = record['SWAPIDPIN'].encode('ascii')
        data += f"|SWAPIDPIN={swapidpin}".encode('ascii')

    if '%UTF8%SWAPIDPART' in record:
        utf8_swapidpart = record['%UTF8%SWAPIDPART'].encode('utf-8')
        data += f"|%UTF8%SWAPIDPART={utf8_swapidpart}".encode('ascii')

    # Write the compiled data to the stream
    stream.write(data + b"\n")  # Assuming newline as a record separator

def handle_ieee_symbol(stream, record):
    """
    Writes an IEEE symbol record to the stream based on its properties.
    """
    # Start the record with its type
    data = b"|RECORD=3"

    # Mandatory fields
    color = int(record.get('COLOR', 255))  # Default color if not specified
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    symbol = int(record['SYMBOL'])
    is_not_accessible = 'T'  # Assuming it's always true as per spec
    scale_factor = int(record['SCALEFACTOR'])
    owner_part_id = int(record['OWNERPARTID'])
    owner_part_display_mode = int(record.get('OWNERPARTDISPLAYMODE', 0))  # Default to 0 if not specified
    current_part_id = int(record.get('CURRENTPARTID', 0))  # Default to 0 if not specified

    data += f"|COLOR={color}|LOCATION.X={location_x}|LOCATION.Y={location_y}|SYMBOL={symbol}|ISNOTACCESIBLE={is_not_accessible}|SCALEFACTOR={scale_factor}|OWNERPARTID={owner_part_id}|OWNERPARTDISPLAYMODE={owner_part_display_mode}|CURRENTPARTID={current_part_id}".encode('ascii')

    # Write the compiled data to the stream
    stream.write(data + b"\n")  # Assuming newline as a record separator

def handle_label(stream, record):
    """
    Writes a label record to the stream based on its properties.
    """
    # Start the record with its type
    data = b"|RECORD=4"

    # Mandatory and optional fields
    owner_index = int(record['OWNERINDEX'])
    is_not_accessible = 'T' if record.get('ISNOTACCESIBLE', 'F') == 'T' else 'F'
    index_in_sheet = int(record['INDEXINSHEET'])
    owner_part_id = int(record['OWNERPARTID'])
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    color = int(record.get('COLOR', 0))  # Default to black if not specified
    is_mirrored = 'T' if record.get('ISMIRRORED', 'F') == 'T' else 'F'
    orientation = int(record.get('ORIENTATION', 0))  # Default to 0 if not specified
    justification = int(record.get('JUSTIFICATION', 0))  # Default to bottom left if not specified
    font_id = int(record.get('FONTID', 1))  # Default to the first font if not specified
    text = record.get('TEXT', '').encode('utf-8')  # Default to empty if not specified
    graphically_locked = 'T' if record.get('GRAPHICALLYLOCKED', 'F') == 'T' else 'F'

    data += f"|OWNERINDEX={owner_index}|ISNOTACCESIBLE={is_not_accessible}|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|LOCATION.X={location_x}|LOCATION.Y={location_y}|COLOR={color}|ISMIRRORED={is_mirrored}|ORIENTATION={orientation}|JUSTIFICATION={justification}|FONTID={font_id}|TEXT={text}|GRAPHICALLYLOCKED={graphically_locked}".encode('ascii')

    # Write the compiled data to the stream
    stream.write(data + b"\n")  # Assuming newline as a record separator

def handle_bezier(stream, record):
    """
    Writes a Bezier curve record to the stream.
    """
    data = b"|RECORD=5"
    owner_index = int(record['OWNERINDEX'])
    is_not_accessible = 'T' if record.get('ISNOTACCESIBLE', 'F') == 'T' else 'F'
    owner_part_id = int(record['OWNERPARTID'])
    line_width = int(record['LINEWIDTH'])
    color = int(record['COLOR'])
    location_count = int(record.get('LOCATIONCOUNT', 4))  # Default to 4 control points if not specified
    
    # Control points data
    control_points = []
    for i in range(1, location_count + 1):
        x = int(record[f'X{i}'])
        y = int(record[f'Y{i}'])
        control_points.append((x, y))

    # Construct the data string for control points
    for i, point in enumerate(control_points):
        data += f"|X{i+1}={point[0]}|Y{i+1}={point[1]}".encode('ascii')

    data += f"|OWNERINDEX={owner_index}|ISNOTACCESIBLE={is_not_accessible}|OWNERPARTID={owner_part_id}|LINEWIDTH={line_width}|COLOR={color}".encode('ascii')

    # Write the compiled data to the stream
    stream.write(data + b"\n")

def handle_polyline(stream, record):
    """
    Writes a polyline record to the stream using the utility function to check for field existence.
    This version of the function assumes no defaults and will only write data if the fields are present in the record.
    """
    data = [b"|RECORD=6"]

    # Append data only if the keys exist
    append_if_present(data, record, 'OWNERINDEX')
    append_if_present(data, record, 'ISNOTACCESIBLE')
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'OWNERPARTDISPLAYMODE')
    append_if_present(data, record, 'LINEWIDTH')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'STARTLINESHAPE')
    append_if_present(data, record, 'ENDLINESHAPE')
    append_if_present(data, record, 'LINESHAPESIZE')

    # Handle location points based on LOCATIONCOUNT
    location_count = int(record.get('LOCATIONCOUNT', 0))
    for i in range(1, location_count + 1):
        append_if_present(data, record, f'X{i}', prefix='|')
        append_if_present(data, record, f'Y{i}', prefix='|')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")



def handle_polygon(stream, record):
    """
    Writes a polygon record to the stream.
    """
    data = [b"|RECORD=7"]

    # Use append_if_present for each field to ensure flexibility
    append_if_present(data, record, 'OWNERINDEX')
    append_if_present(data, record, 'ISNOTACCESIBLE', transform=lambda x: 'T' if x == 'T' else 'F')
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'OWNERPARTDISPLAYMODE')
    append_if_present(data, record, 'LINEWIDTH')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'AREACOLOR')
    append_if_present(data, record, 'ISSOLID', transform=lambda x: 'T' if x == 'T' else 'F')
    append_if_present(data, record, 'IGNOREONLOAD', transform=lambda x: 'T' if x == 'T' else 'F')

    # Add properties for each location point if LOCATIONCOUNT is present
    if 'LOCATIONCOUNT' in record:
        location_count = int(record['LOCATIONCOUNT'])
        for i in range(1, location_count + 1):
            append_if_present(data, record, f'X{i}')
            append_if_present(data, record, f'Y{i}')

    # Extra points if specified
    if 'EXTRALOCATIONCOUNT' in record:
        extra_location_count = int(record['EXTRALOCATIONCOUNT'])
        for i in range(1, extra_location_count + 1):
            append_if_present(data, record, f'EX{i}')
            append_if_present(data, record, f'EY{i}')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")



def handle_ellipse(stream, record):
    """
    Writes an ellipse record to the stream.
    """
    data = [b"|RECORD=8"]

    # Use append_if_present for each field to ensure flexibility
    append_if_present(data, record, 'OWNERINDEX')
    append_if_present(data, record, 'ISNOTACCESIBLE', transform=lambda x: 'T' if x == 'T' else 'F')
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'RADIUS')
    append_if_present(data, record, 'SECONDARYRADIUS')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'AREACOLOR')
    append_if_present(data, record, 'ISSOLID', transform=lambda x: 'T' if x == 'T' else 'F')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'LINEWIDTH')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


# Further handlers for Piechart, Round Rectangle, Elliptical Arc, and Arc can be defined similarly.
def handle_piechart(stream, record):
    """
    Writes a piechart record to the stream, assuming basic properties similar to other visual objects.
    """
    data = b"|RECORD=9"
    # Here you would add the properties relevant to a Piechart, which are not specified.
    # For example:
    # location_x = int(record['LOCATION.X'])
    # location_y = int(record['LOCATION.Y'])
    # color = int(record['COLOR'])
    # data += f"|LOCATION.X={location_x}|LOCATION.Y={location_y}|COLOR={color}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_round_rectangle(stream, record):
    """
    Writes a round rectangle record to the stream.
    """
    data = b"|RECORD=10"
    owner_index = int(record['OWNERINDEX'])
    is_not_accessible = record.get('ISNOTACCESIBLE', 'T')  # Default to True if not specified
    index_in_sheet = int(record['INDEXINSHEET'])
    owner_part_id = int(record['OWNERPARTID'])
    owner_part_display_mode = int(record.get('OWNERPARTDISPLAYMODE', 0))  # Default to 0 if not specified
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    corner_x = int(record['CORNER.X'])
    corner_y = int(record['CORNER.Y'])
    corner_x_radius = int(record.get('CORNERXRADIUS', 0))
    corner_y_radius = int(record.get('CORNERYRADIUS', 0))
    line_width = int(record['LINEWIDTH'])
    color = int(record['COLOR'])
    area_color = int(record.get('AREACOLOR', 0))  # Default to 0 if not specified
    is_solid = record.get('ISSOLID', 'F')  # Default to False if not specified

    # Compile the data string
    data += f"|OWNERINDEX={owner_index}|ISNOTACCESIBLE={is_not_accessible}|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|OWNERPARTDISPLAYMODE={owner_part_display_mode}|LOCATION.X={location_x}|LOCATION.Y={location_y}|CORNER.X={corner_x}|CORNER.Y={corner_y}|CORNERXRADIUS={corner_x_radius}|CORNERYRADIUS={corner_y_radius}|LINEWIDTH={line_width}|COLOR={color}|AREACOLOR={area_color}|ISSOLID={is_solid}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_sheet_entry(stream, record):
    """
    Writes a sheet entry record to the stream.
    Sheet entries define ports on a sheet symbol in schematic documents.
    """
    data = b"|RECORD=16"
    area_color = int(record['AREACOLOR'])
    arrow_kind = record['ARROWKIND']
    color = int(record['COLOR'])
    distance_from_top = int(record['DISTANCEFROMTOP'])
    distance_from_top_frac = int(record['DISTANCEFROMTOP_FRAC1'])
    name = record['NAME'].encode('ascii')
    owner_part_id = int(record['OWNERPARTID'])
    owner_index = int(record['OWNERINDEX'])
    text_color = int(record['TEXTCOLOR'])
    text_font_id = int(record['TEXTFONTID'])
    text_style = record['TEXTSTYLE'].encode('ascii')
    index_in_sheet = int(record['INDEXINSHEET'])
    harness_type = record['HARNESSTYPE'].encode('ascii')
    side = int(record['SIDE'])
    style = int(record['STYLE'])
    io_type = int(record['IOTYPE'])

    # Compile the data string
    data += f"|AREACOLOR={area_color}|ARROWKIND={arrow_kind}|COLOR={color}|DISTANCEFROMTOP={distance_from_top}|DISTANCEFROMTOP_FRAC1={distance_from_top_frac}|NAME={name}|OWNERPARTID={owner_part_id}|OWNERINDEX={owner_index}|TEXTCOLOR={text_color}|TEXTFONTID={text_font_id}|TEXTSTYLE={text_style}|INDEXINSHEET={index_in_sheet}|HARNESSTYPE={harness_type}|SIDE={side}|STYLE={style}|IOTYPE={io_type}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")


def handle_power_port(stream, record):
    """
    Writes a power port record to the stream.
    Power ports represent electrical connections to power rails, ground, or other common points.
    This version of the function assumes no defaults and will only write data if the fields are present in the record.
    """
    data = [b"|RECORD=17"]

    # Append data only if the keys exist
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'STYLE')
    append_if_present(data, record, 'SHOWNETNAME', transform=lambda x: 'T' if x else 'F')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'ORIENTATION')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'TEXT', transform=lambda x: x.encode('ascii'))
    append_if_present(data, record, 'ISCROSSSHEETCONNECTOR', transform=lambda x: 'T' if x else 'F')
    append_if_present(data, record, 'UNIQUEID', transform=lambda x: x.encode('ascii'))
    append_if_present(data, record, 'FONTID')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_port(stream, record):
    """
    Writes a port record to the stream.
    Ports represent connections and are typically labeled with information like signal names or electrical characteristics.
    """
    data = [b"|RECORD=18"]

    # Using append_if_present for each field
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'STYLE')
    append_if_present(data, record, 'IOTYPE')
    append_if_present(data, record, 'ALIGNMENT')
    append_if_present(data, record, 'WIDTH')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'AREACOLOR')
    append_if_present(data, record, 'TEXTCOLOR')
    append_if_present(data, record, 'NAME', transform=lambda x: x.encode('ascii'))
    append_if_present(data, record, 'UNIQUEID', transform=lambda x: x.encode('ascii'))
    append_if_present(data, record, 'FONTID')
    append_if_present(data, record, 'HEIGHT')
    append_if_present(data, record, 'HARNESSTYPE', transform=lambda x: x.encode('ascii'))

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")



def handle_elliptical_arc(stream, record):
    """
    Writes an elliptical arc record to the stream.
    """
    data = b"|RECORD=11"
    owner_index = int(record['OWNERINDEX'])
    is_not_accessible = 'T'
    index_in_sheet = int(record['INDEXINSHEET'])
    owner_part_id = int(record['OWNERPARTID'])
    owner_part_display_mode = int(record['OWNERPARTDISPLAYMODE'])
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    radius = int(record['RADIUS'])
    radius_frac = int(record.get('RADIUS_FRAC', 0))
    secondary_radius = int(record['SECONDARYRADIUS'])
    secondary_radius_frac = int(record.get('SECONDARYRADIUS_FRAC', 0))
    line_width = int(record['LINEWIDTH'])
    start_angle = float(record['STARTANGLE'])
    end_angle = float(record['ENDANGLE'])
    color = int(record['COLOR'])

    # Compile the data string
    data += f"|OWNERINDEX={owner_index}|ISNOTACCESIBLE={is_not_accessible}|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|OWNERPARTDISPLAYMODE={owner_part_display_mode}|LOCATION.X={location_x}|LOCATION.Y={location_y}|RADIUS={radius}|RADIUS_FRAC={radius_frac}|SECONDARYRADIUS={secondary_radius}|SECONDARYRADIUS_FRAC={secondary_radius_frac}|LINEWIDTH={line_width}|STARTANGLE={start_angle}|ENDANGLE={end_angle}|COLOR={color}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_arc(stream, record):
    """
    Writes a circular arc record to the stream using the utility function to check for field existence.
    """
    data = [b"|RECORD=12"]

    # Using the utility function to conditionally append data
    append_if_present(data, record, 'OWNERINDEX')
    append_if_present(data, record, 'ISNOTACCESIBLE', format_string='', prefix='')
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'OWNERPARTDISPLAYMODE')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'RADIUS')
    append_if_present(data, record, 'RADIUS_FRAC')
    append_if_present(data, record, 'LINEWIDTH')
    append_if_present(data, record, 'STARTANGLE', format_string='', prefix='')
    append_if_present(data, record, 'ENDANGLE', format_string='', prefix='')
    append_if_present(data, record, 'COLOR')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_line(stream, record):
    """
    Writes a line record to the stream using the utility function to check for field existence.
    """
    data = [b"|RECORD=13"]

    # Using the utility function to conditionally append data
    append_if_present(data, record, 'OWNERINDEX')
    append_if_present(data, record, 'ISNOTACCESIBLE', format_string='', prefix='')
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'OWNERPARTDISPLAYMODE')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'CORNER.X')
    append_if_present(data, record, 'CORNER.Y')
    append_if_present(data, record, 'LINEWIDTH')
    append_if_present(data, record, 'COLOR')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")



def handle_rectangle(stream, record):
    """
    Writes a rectangle record to the stream.
    """
    data = b"|RECORD=14"
    owner_index = int(record['OWNERINDEX'])
    is_not_accessible = 'T'
    index_in_sheet = int(record['INDEXINSHEET'])
    owner_part_id = int(record['OWNERPARTID'])
    owner_part_display_mode = int(record.get('OWNERPARTDISPLAYMODE', ''))
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    corner_x = int(record['CORNER.X'])
    corner_y = int(record['CORNER.Y'])
    line_width = int(record['LINEWIDTH'])
    color = int(record['COLOR'])
    area_color = int(record['AREACOLOR'])
    is_solid = 'T' if record.get('ISSOLID', 'F') == 'T' else 'F'
    transparent = 'T' if record.get('TRANSPARENT', 'F') == 'T' else 'F'

    # Compile the data string
    data += f"|OWNERINDEX={owner_index}|ISNOTACCESIBLE={is_not_accessible}|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|OWNERPARTDISPLAYMODE={owner_part_display_mode}|LOCATION.X={location_x}|LOCATION.Y={location_y}|CORNER.X={corner_x}|CORNER.Y={corner_y}|LINEWIDTH={line_width}|COLOR={color}|AREACOLOR={area_color}|ISSOLID={is_solid}|TRANSPARENT={transparent}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_sheet_symbol(stream, record):
    """
    Writes a sheet symbol record to the stream.
    """
    data = b"|RECORD=15"
    index_in_sheet = int(record['INDEXINSHEET'])
    owner_part_id = -1  # Always -1 for sheet symbols
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    x_size = int(record['XSIZE'])
    y_size = int(record['YSIZE'])
    color = int(record['COLOR'])
    area_color = int(record['AREACOLOR'])
    is_solid = 'T'
    unique_id = record['UNIQUEID']
    symbol_type = record.get('SYMBOLTYPE', 'Normal')

    # Compile the data string
    data += f"|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|LOCATION.X={location_x}|LOCATION.Y={location_y}|XSIZE={x_size}|YSIZE={y_size}|COLOR={color}|AREACOLOR={area_color}|ISSOLID={is_solid}|UNIQUEID={unique_id}|SYMBOLTYPE={symbol_type}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_no_erc(stream, record):
    """
    Writes a No ERC (No Electrical Rule Check) record to the stream.
    This record typically indicates that a certain connection does not require electrical rule checking.
    """
    data = [b"|RECORD=22"]

    # Using the append_if_present function to conditionally append data
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'ISACTIVE', transform=lambda x: 'T' if x == 'True' else 'F')
    append_if_present(data, record, 'ORIENTATION')
    append_if_present(data, record, 'SUPPRESSALL', transform=lambda x: 'T' if x == 'True' else 'F')
    append_if_present(data, record, 'SYMBOL', transform=lambda x: x.encode('ascii'))
    append_if_present(data, record, 'UNIQUEID', transform=lambda x: x.encode('ascii'))

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_net_label(stream, record):
    """
    Writes a Net Label record to the stream.
    Net labels are used to identify connections in schematic diagrams.
    """
    data = [b"|RECORD=25"]

    # Using append_if_present for each field without defaults
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'FONTID')
    append_if_present(data, record, 'TEXT', transform=lambda x: x.encode('ascii'))
    append_if_present(data, record, 'ORIENTATION')
    append_if_present(data, record, 'UNIQUEID', transform=lambda x: x.encode('ascii'))

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_junction(stream, record):
    """
    Writes a Junction record to the stream.
    Junctions are typically used to denote electrical connections between wires or pins.
    """
    data = [b"|RECORD=29"]

    # Use append_if_present for each field to ensure flexibility
    append_if_present(data, record, 'INDEXINSHEET')
    append_if_present(data, record, 'LOCKED', transform=lambda x: 'T' if x else 'F')
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'COLOR')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_image(stream, record):
    """
    Writes an Image record to the stream.
    Images can be embedded or linked within the schematic file.
    """
    data = b"|RECORD=30"
    owner_index = int(record['OWNERINDEX'])
    index_in_sheet = int(record['INDEXINSHEET'])
    owner_part_id = int(record.get('OWNERPARTID', -1))  # Default to -1 if not provided
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    corner_x = int(record['CORNER.X'])
    corner_y = int(record['CORNER.Y'])
    embed_image = 'T' if record['EMBEDIMAGE'] else 'F'  # Assuming boolean is provided as True/False
    file_name = record['FILENAME'].encode('ascii')  # Assume ASCII, but consider cp1252 for non-ASCII
    keep_aspect = 'T' if record['KEEPASPECT'] else 'F'  # Assuming boolean is provided as True/False

    # Compile the data string
    data += f"|OWNERINDEX={owner_index}|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|LOCATION.X={location_x}|LOCATION.Y={location_y}|CORNER.X={corner_x}|CORNER.Y={corner_y}|EMBEDIMAGE={embed_image}|FILENAME={file_name}|KEEPASPECT={keep_aspect}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_sheet(stream, record):
    """
    Writes a Sheet record to the stream. This defines the overall settings for a schematic sheet.
    """
    data = b"|RECORD=31"
    font_id_count = int(record['FONTIDCOUNT'])
    size_n = int(record['SIZEn'])
    italic_n = 'T' if record['ITALICn'] else 'F'
    bold_n = 'T' if record['BOLDn'] else 'F'
    underline_n = 'T' if record['UNDERLINEn'] else 'F'
    rotation_n = int(record['ROTATIONn'])
    font_name_n = record['FONTNAMEn'].encode('utf-8')
    use_mbc = 'T' if record['USEMBCS'] else 'F'
    is_boc = 'T' if record['ISBOC'] else 'F'
    hotspot_grid_on = 'T' if record['HOTSPOTGRIDON'] else 'F'
    hotspot_grid_size = int(record['HOTSPOTGRIDSIZE'])
    sheet_style = int(record['SHEETSTYLE'])
    system_font = int(record['SYSTEMFONT'])
    border_on = 'T' if record['BORDERON'] else 'F'
    title_block_on = 'T' if record['TITLEBLOCKON'] else 'F'
    sheet_number_space_size = int(record['SHEETNUMBERSPACESIZE'])
    area_color = int(record['AREACOLOR'])
    snap_grid_on = 'T' if record['SNAPGRIDON'] else 'F'
    snap_grid_size = int(record['SNAPGRIDSIZE'])
    visible_grid_on = 'T' if record['VISIBLEGRIDON'] else 'F'
    visible_grid_size = int(record['VISIBLEGRIDSIZE'])
    custom_x = int(record['CUSTOMX'])
    custom_y = int(record['CUSTOMY'])
    use_custom_sheet = 'T' if record['USECUSTOMSHEET'] else 'F'
    workspace_orientation = int(record['WORKSPACEORIENTATION'])
    custom_x_zones = int(record['CUSTOMXZONES'])
    custom_y_zones = int(record['CUSTOMYZONES'])
    custom_margin_width = int(record['CUSTOMMARGINWIDTH'])
    display_unit = int(record['DISPLAY_UNIT'])
    reference_zones_on = 'T' if record['REFERENCEZONESON'] else 'F'
    show_template_graphics = 'T' if record['SHOWTEMPLATEGRAPHICS'] else 'F'
    template_filename = record['TEMPLATEFILENAME'].encode('utf-8')

    # Compile the data string
    data += f"|FONTIDCOUNT={font_id_count}|SIZEn={size_n}|ITALICn={italic_n}|BOLDn={bold_n}|UNDERLINEn={underline_n}|ROTATIONn={rotation_n}|FONTNAMEn={font_name_n}|USEMBCS={use_mbc}|ISBOC={is_boc}|HOTSPOTGRIDON={hotspot_grid_on}|HOTSPOTGRIDSIZE={hotspot_grid_size}|SHEETSTYLE={sheet_style}|SYSTEMFONT={system_font}|BORDERON={border_on}|TITLEBLOCKON={title_block_on}|SHEETNUMBERSPACESIZE={sheet_number_space_size}|AREACOLOR={area_color}|SNAPGRIDON={snap_grid_on}|SNAPGRIDSIZE={snap_grid_size}|VISIBLEGRIDON={visible_grid_on}|VISIBLEGRIDSIZE={visible_grid_size}|CUSTOMX={custom_x}|CUSTOMY={custom_y}|USECUSTOMSHEET={use_custom_sheet}|WORKSPACEORIENTATION={workspace_orientation}|CUSTOMXZONES={custom_x_zones}|CUSTOMYZONES={custom_y_zones}|CUSTOMMARGINWIDTH={custom_margin_width}|DISPLAY_UNIT={display_unit}|REFERENCEZONESON={reference_zones_on}|SHOWTEMPLATEGRAPHICS={show_template_graphics}|TEMPLATEFILENAME={template_filename}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_designator(stream, record):
    """
    Writes a Designator record to the stream. This involves component identification labels on schematics.
    """
    data = [b"|RECORD=34"]

    # Conditional inclusion of each field
    if 'OWNERINDEX' in record:
        owner_index = struct.pack('<I', int(record['OWNERINDEX']))
        data.append(b"|OWNERINDEX=" + owner_index)
    if 'INDEXINSHEET' in record:
        index_in_sheet = struct.pack('<I', int(record['INDEXINSHEET']))
        data.append(b"|INDEXINSHEET=" + index_in_sheet)
    if 'OWNERPARTID' in record:
        owner_part_id = struct.pack('<I', int(record['OWNERPARTID']))
        data.append(b"|OWNERPARTID=" + owner_part_id)
    if 'LOCATION.X' in record:
        location_x = struct.pack('<I', int(record['LOCATION.X']))
        data.append(b"|LOCATION.X=" + location_x)
    if 'LOCATION.Y' in record:
        location_y = struct.pack('<I', int(record['LOCATION.Y']))
        data.append(b"|LOCATION.Y=" + location_y)
    if 'COLOR' in record:
        color = struct.pack('<I', int(record['COLOR']))
        data.append(b"|COLOR=" + color)
    if 'FONTID' in record:
        font_id = struct.pack('<I', int(record['FONTID']))
        data.append(b"|FONTID=" + font_id)
    if 'TEXT' in record:
        text = record['TEXT'].encode('utf-8')
        data.append(b"|TEXT=" + text)
    if 'NAME' in record:
        name = record['NAME'].encode('utf-8')
        data.append(b"|NAME=" + name)
    if 'READONLYSTATE' in record:
        readonly_state = struct.pack('<I', int(record['READONLYSTATE']))
        data.append(b"|READONLYSTATE=" + readonly_state)
    if 'ISMIRRORED' in record:
        is_mirrored = b'T' if record['ISMIRRORED'] == 'T' else b'F'
        data.append(b"|ISMIRRORED=" + is_mirrored)
    if 'ORIENTATION' in record:
        orientation = struct.pack('<I', int(record['ORIENTATION']))
        data.append(b"|ORIENTATION=" + orientation)
    if 'ISHIDDEN' in record:
        is_hidden = b'T' if record['ISHIDDEN'] == 'T' else b'F'
        data.append(b"|ISHIDDEN=" + is_hidden)
    if 'UNIQUEID' in record:
        unique_id = record['UNIQUEID'].encode('utf-8')
        data.append(b"|UNIQUEID=" + unique_id)
    if 'OVERRIDENOTAUTOPOSITION' in record:
        override_not_auto_position = b'T' if record['OVERRIDENOTAUTOPOSITION'] == 'T' else b'F'
        data.append(b"|OVERRIDENOTAUTOPOSITION=" + override_not_auto_position)

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_bus_entry(stream, record):
    """
    Writes a Bus Entry record to the stream. This record represents bus entry points on schematics.
    """
    data = b"|RECORD=37"
    color = int(record.get('COLOR', 8388608))  # Default dark grey color.
    corner_x = int(record['CORNER.X'])
    corner_y = int(record['CORNER.Y'])
    line_width = int(record['LINEWIDTH'])
    location_x = int(record['LOCATION.X'])
    location_y = int(record['LOCATION.Y'])
    owner_part_id = int(record.get('OWNERPARTID', -1))

    # Compile the data string
    data += f"|COLOR={color}|CORNER.X={corner_x}|CORNER.Y={corner_y}|LINEWIDTH={line_width}|LOCATION.X={location_x}|LOCATION.Y={location_y}|OWNERPARTID={owner_part_id}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_template(stream, record):
    """
    Writes a Template record to the stream. This record manages custom template settings such as title blocks.
    """
    data = b"|RECORD=39"
    is_not_accessible = 'T' if record.get('ISNOTACCESIBLE', True) else 'F'  # Default to True (Not accessible).
    owner_part_id = int(record.get('OWNERPARTID', -1))
    filename = record['FILENAME'].encode('utf-8')

    # Compile the data string
    data += f"|ISNOTACCESIBLE={is_not_accessible}|OWNERPARTID={owner_part_id}|FILENAME={filename}".encode('ascii')

    # Write to stream
    stream.write(data + b"\n")

def handle_parameter(stream, record):
    data = [b"|RECORD=41"]

    # Conditional inclusion of each field
    if 'INDEXINSHEET' in record:
        index_in_sheet = struct.pack('<I', int(record['INDEXINSHEET']))
        data.append(b"|INDEXINSHEET=" + index_in_sheet)
    if 'OWNERINDEX' in record:
        owner_index = struct.pack('<I', int(record['OWNERINDEX']))
        data.append(b"|OWNERINDEX=" + owner_index)
    if 'OWNERPARTID' in record:
        owner_part_id = struct.pack('<I', int(record['OWNERPARTID']))
        data.append(b"|OWNERPARTID=" + owner_part_id)
    if 'LOCATION.X' in record:
        location_x = struct.pack('<I', int(record['LOCATION.X']))
        data.append(b"|LOCATION.X=" + location_x)
    if 'LOCATION.Y' in record:
        location_y = struct.pack('<I', int(record['LOCATION.Y']))
        data.append(b"|LOCATION.Y=" + location_y)
    if 'ORIENTATION' in record:
        orientation = struct.pack('<I', int(record['ORIENTATION']))
        data.append(b"|ORIENTATION=" + orientation)
    if 'COLOR' in record:
        color = struct.pack('<I', int(record['COLOR']))
        data.append(b"|COLOR=" + color)
    if 'FONTID' in record:
        font_id = struct.pack('<I', int(record['FONTID']))
        data.append(b"|FONTID=" + font_id)
    if 'ISHIDDEN' in record:
        is_hidden = b'T' if record['ISHIDDEN'] == 'T' else b'F'
        data.append(b"|ISHIDDEN=" + is_hidden)
    if 'TEXT' in record:
        text = record['TEXT'].encode('utf-8')
        data.append(b"|TEXT=" + text)
    if 'NAME' in record:
        name = record['NAME'].encode('utf-8')
        data.append(b"|NAME=" + name)
    if '%UTF8%TEXT' in record:
        utf8_text = record['%UTF8%TEXT'].encode('utf-8')
        data.append(b"|UTF8%TEXT=" + utf8_text)
    if '%UTF8%NAME' in record:
        utf8_name = record['%UTF8%NAME'].encode('utf-8')
        data.append(b"|UTF8%NAME=" + utf8_name)
    if 'READONLYSTATE' in record:
        read_only_state = struct.pack('<I', int(record['READONLYSTATE']))
        data.append(b"|READONLYSTATE=" + read_only_state)
    if 'UNIQUEID' in record:
        unique_id = record['UNIQUEID'].encode('utf-8')
        data.append(b"|UNIQUEID=" + unique_id)
    if 'ISMIRRORED' in record:
        is_mirrored = b'T' if record['ISMIRRORED'] == 'T' else b'F'
        data.append(b"|ISMIRRORED=" + is_mirrored)
    if 'NOTAUTOPOSITION' in record:
        not_auto_position = b'T' if record['NOTAUTOPOSITION'] == 'T' else b'F'
        data.append(b"|NOTAUTOPOSITION=" + not_auto_position)
    if 'SHOWNAME' in record:
        show_name = b'T' if record['SHOWNAME'] == 'T' else b'F'
        data.append(b"|SHOWNAME=" + show_name)

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")
    print(final_data)




def handle_warning_sign(stream, record):
    """
    Writes a Warning Sign record to the stream. This record is used for marking important circuit characteristics.
    """
    data = [b"|RECORD=43"]

    # Using append_if_present for each field
    append_if_present(data, record, 'OWNERPARTID')
    append_if_present(data, record, 'LOCATION.X')
    append_if_present(data, record, 'LOCATION.Y')
    append_if_present(data, record, 'COLOR')
    append_if_present(data, record, 'NAME', transform=lambda x: x.encode('utf-8'))
    append_if_present(data, record, 'ORIENTATION')

    # Combine all parts into a single byte string and write to stream
    final_data = b''.join(data)
    stream.write(final_data + b"\n")


def handle_implementation(stream, record):
    """
    Writes an Implementation record to the stream, detailing specific model files linked to components.
    """
    data = b"|RECORD=45"
    owner_index = int(record['OWNERINDEX'])
    model_name = record.get('MODELNAME', '').encode('utf-8')
    model_type = record.get('MODELTYPE', '').encode('utf-8')
    data_file_count = int(record.get('DATAFILECOUNT', 0))
    data_file_entity = record.get('MODELDATAFILEENTITY0', '').encode('utf-8')
    data_file_kind = record.get('MODELDATAFILEKIND0', '').encode('utf-8')
    is_current = 'T' if record.get('ISCURRENT', False) else 'F'

    # Compile the data string with optional parameters
    data += (f"|OWNERINDEX={owner_index}|MODELNAME={model_name}|MODELTYPE={model_type}"
             f"|DATAFILECOUNT={data_file_count}|MODELDATAFILEENTITY0={data_file_entity}"
             f"|MODELDATAFILEKIND0={data_file_kind}|ISCURRENT={is_current}").encode('utf-8')

    # Additional optional fields
    if 'DESCRIPTION' in record:
        description = record['DESCRIPTION'].encode('utf-8')
        data += f"|DESCRIPTION={description}".encode('utf-8')
    if 'USECOMPONENTLIBRARY' in record:
        use_component_library = 'T' if record['USECOMPONENTLIBRARY'] else 'F'
        data += f"|USECOMPONENTLIBRARY={use_component_library}".encode('utf-8')

    # Write to stream
    stream.write(data + b"\n")

def handle_record_46(stream, record):
    """
    Handles a specific sub-model or variant linked to an implementation (Record 45).
    This function simply records the ownership to a parent implementation.
    """
    owner_index = int(record['OWNERINDEX'])
    data = f"|RECORD=46|OWNERINDEX={owner_index}".encode('utf-8')
    stream.write(data + b"\n")

def handle_record_47(stream, record):
    """
    A generic handler for Record 47, which may need more specific information to fully implement.
    """
    owner_index = int(record['OWNERINDEX'])
    index_in_sheet = int(record['INDEXINSHEET'])
    # Add other properties as needed
    data = f"|RECORD=47|OWNERINDEX={owner_index}|INDEXINSHEET={index_in_sheet}".encode('utf-8')
    stream.write(data + b"\n")

def handle_record_48(stream, record):
    """
    Another child type of Record 45, often used for linking additional data or sub-components.
    """
    owner_index = int(record['OWNERINDEX'])
    data = f"|RECORD=48|OWNERINDEX={owner_index}".encode('utf-8')
    stream.write(data + b"\n")

def handle_record_215_to_218(stream, record):
    """
    Handles records typically found in the Additional Stream, relating to additional sheet properties or annotations.
    """
    record_type = record['RECORD']
    color = record.get('COLOR', 0)
    index_in_sheet = record.get('INDEXINSHEET', -1)
    line_width = record.get('LINEWIDTH', 1)
    location_count = record.get('LOCATIONCOUNT', 0)
    owner_part_id = record.get('OWNERPARTID', -1)
    coordinates = f"|X1={record.get('X1',0)}|X2={record.get('X2',0)}|Y1={record.get('Y1',0)}|Y2={record.get('Y2',0)}"

    data = f"|RECORD={record_type}|COLOR={color}|INDEXINSHEET={index_in_sheet}|LINEWIDTH={line_width}|LOCATIONCOUNT={location_count}|OWNERPARTID={owner_part_id}{coordinates}".encode('utf-8')
    stream.write(data + b"\n")

def handle_hyperlink(stream, record):
    """
    Handles a hyperlink record, possibly linking to external documents or URLs.
    """
    # Assume additional fields such as URL or description might be provided in the record dictionary
    data = b"|RECORD=226"
    # Example additional fields (not provided in your data)
    url = record.get('URL', '').encode('utf-8')
    description = record.get('DESCRIPTION', '').encode('utf-8')
    
    data += f"|URL={url}|DESCRIPTION={description}".encode('utf-8')
    stream.write(data + b"\n")

def json_to_schdoc(json_path, schdoc_path):
    logging.info(f"Reading JSON data from: {json_path}")
    with open(json_path, 'r') as file:
        data = json.load(file)
    
    logging.info(f"Attempting to create or open OLE file at: {schdoc_path}")
    with olefile.OleFileIO(schdoc_path, write_mode=True) as ole:
        # Check and process records
        print(ole.listdir())
        #assert True == False
        if 'records' in data:
            logging.info("Creating or opening FileHeader stream...")
            create_fileheader_stream(ole, data['records'])
        else:
            logging.warning("No 'records' key found in JSON data.")
        
        logging.info("OLE file changes will be auto-saved on close.")



    # Closing the OLE file should be outside the with block if not using with as context manager
    ole.close()



def create_fileheader_stream(ole, data):
    """
    Creates the FileHeader stream with appropriate object records in the OLE file.
    """
    if not ole.exists('FileHeader'):
        assert True == False
        #ole.create_stream('RECORD')
    with ole.openstream('FileHeader') as stream:
        for record in data:
            # Attempt to get and convert the record type to integer
            record_type_str = record.get('RECORD')
            if record_type_str is None:
                logging.warning("Record type is missing in this record, skipping...")
                continue

            try:
                record_type = int(record_type_str)  # Convert record type to integer
            except ValueError:
                logging.error(f"Invalid record type {record_type_str} encountered, skipping...")
                continue

            handler = record_handlers.get(record_type)
            if handler:
                #logging.info(f"Processing record type: {record_type}")
                handler(stream, record)
            else:
                logging.warning(f"Unhandled or missing record type: {record_type}")
        

        
        #with open("out.SchDoc", "wb") as output_file:
        #    output_file.write(stream.read())
            #output_file.write(stream.openstream("FileHeader").read())
            #output_file.write(stream.openstream("Storage").read())

        logging.info("Data written to FileHeader stream.")


def write_pin_record(stream, record):
    """
    Writes a pin record to the stream.
    """
    try:
        # Ensure all required fields are present
        owneridx = struct.pack('<I', int(record['OWNERINDEX']))
        locx = struct.pack('<I', int(record['LOCATION.X']))
        locy = struct.pack('<I', int(record['LOCATION.Y']))

        # Build the binary data for the pin record
        data = b'|RECORD=2' + b'|OWNERINDEX=' + owneridx + b'|LOCATION.X=' + locx + b'|LOCATION.Y=' + locy + b'|'
        stream.write(data)
    except KeyError as e:
        print(f"Missing expected property: {str(e)}")
    except ValueError as e:
        print(f"Data format error: {str(e)}")

def write_line_record(stream, record):
    try:
        start_x = struct.pack('<I', int(record['LOCATION.X']))
        start_y = struct.pack('<I', int(record['LOCATION.Y']))
        end_x = struct.pack('<I', int(record['CORNER.X']))
        end_y = struct.pack('<I', int(record['CORNER.Y']))
        line_width = struct.pack('<I', int(record.get('LINEWIDTH', 1)))  # Default width

        data = (b'|RECORD=13' + b'|START.X=' + start_x + b'|START.Y=' + start_y +
                b'|END.X=' + end_x + b'|END.Y=' + end_y + b'|LINEWIDTH=' + line_width + b'|')
        stream.write(data)
    except KeyError as e:
        print(f"Missing expected property for line: {str(e)}")

def handle_rectangle(stream, record):
    data = [b'|RECORD=14']

    # Conditional inclusion of each field
    if 'LOCATION.X' in record:
        loc_x = struct.pack('<I', int(record['LOCATION.X']))
        data.append(b'|LOCATION.X=' + loc_x)
    if 'LOCATION.Y' in record:
        loc_y = struct.pack('<I', int(record['LOCATION.Y']))
        data.append(b'|LOCATION.Y=' + loc_y)
    if 'CORNER.X' in record:
        corner_x = struct.pack('<I', int(record['CORNER.X']))
        data.append(b'|CORNER.X=' + corner_x)
    if 'CORNER.Y' in record:
        corner_y = struct.pack('<I', int(record['CORNER.Y']))
        data.append(b'|CORNER.Y=' + corner_y)
    if 'LINEWIDTH' in record:
        line_width = struct.pack('<I', int(record['LINEWIDTH']))
        data.append(b'|LINEWIDTH=' + line_width)
    if 'OWNERINDEX' in record:
        owner_index = struct.pack('<I', int(record['OWNERINDEX']))
        data.append(b'|OWNERINDEX=' + owner_index)

    # Combine all parts into a single byte string
    final_data = b''.join(data)

    # Write to stream
    try:
        stream.write(final_data + b'\n')
    except Exception as e:
        logging.error(f"Error writing rectangle record to stream: {e}")


def handle_bus(stream, record):
    """
    Handles the Bus record, which defines a polyline representing a bus in the schematic.
    """
    index_in_sheet = int(record.get('INDEXINSHEET', -1))  # Default to -1 if not specified
    owner_part_id = int(record.get('OWNERPARTID', -1))  # Default to -1 if not specified
    line_width = int(record.get('LINEWIDTH', 1))  # Default line width if not specified
    color = int(record.get('COLOR', 0))  # Default color (black) if not specified
    location_count = int(record.get('LOCATIONCOUNT', 0))  # Default to 0 if not specified
    
    data = f"|RECORD=26|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|LINEWIDTH={line_width}|COLOR={color}|LOCATIONCOUNT={location_count}".encode('utf-8')
    stream.write(data)
    
    # Process coordinate pairs
    for i in range(location_count):
        x = int(record.get(f'X{i}', 0))  # Default to 0 if not specified
        y = int(record.get(f'Y{i}', 0))  # Default to 0 if not specified
        data = f"|X{i}={x}|Y{i}={y}".encode('utf-8')
        stream.write(data)
    
    stream.write(b'\n')  # Add a newline to separate records

def handle_wire(stream, record):
    """
    Handles the Wire record, which defines a polyline representing a wire connection in the schematic.
    """
    index_in_sheet = int(record.get('INDEXINSHEET', -1))  # Default to -1 if not specified
    owner_part_id = int(record.get('OWNERPARTID', -1))  # Default to -1 if not specified
    line_width = int(record.get('LINEWIDTH', 1))  # Default line width if not specified
    color = int(record.get('COLOR', 0))  # Default color (black) if not specified
    location_count = int(record.get('LOCATIONCOUNT', 0))  # Default to 0 if not specified
    unique_id = record.get('UNIQUEID', '')  # Default to empty string if not specified
    
    data = f"|RECORD=27|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|LINEWIDTH={line_width}|COLOR={color}|LOCATIONCOUNT={location_count}|UNIQUEID={unique_id}".encode('utf-8')
    stream.write(data)
    
    # Process coordinate pairs
    for i in range(location_count):
        x = int(record.get(f'X{i}', 0))  # Default to 0 if not specified
        y = int(record.get(f'Y{i}', 0))  # Default to 0 if not specified
        data = f"|X{i}={x}|Y{i}={y}".encode('utf-8')
        stream.write(data)
    
    stream.write(b'\n')  # Add a newline to separate records

def handle_text_frame(stream, record):
    """
    Handles the Text Frame record, which defines a text box possibly with word wrapping and other properties.
    """
    owner_part_id = int(record.get('OWNERPARTID', -1))  # Default to -1 if not specified
    location_x = int(record.get('LOCATION.X', 0))  # Default to 0 if not specified
    location_y = int(record.get('LOCATION.Y', 0))  # Default to 0 if not specified
    corner_x = int(record.get('CORNER.X', 0))  # Default to 0 if not specified
    corner_y = int(record.get('CORNER.Y', 0))  # Default to 0 if not specified
    area_color = int(record.get('AREACOLOR', 16777215))  # Default white if not specified
    font_id = int(record.get('FONTID', 1))  # Default to 1 if not specified
    alignment = int(record.get('ALIGNMENT', 1))  # Default alignment if not specified
    word_wrap = record.get('WORDWRAP', 'T')  # Default to true if not specified
    color = int(record.get('COLOR', 0))  # Default color (black) if not specified
    is_solid = record.get('ISSOLID', 'F')  # Default to false if not specified
    clip_to_rect = record.get('CLIPTORECT', 'F')  # Default to false if not specified
    is_not_accessible = record.get('ISNOTACCESIBLE', 'F')  # Default to false if not specified
    orientation = int(record.get('ORIENTATION', 0))  # Default orientation if not specified
    text_margin_frac = int(record.get('TEXTMARGIN_FRAC', 0))  # Default to 0 if not specified
    index_in_sheet = int(record.get('INDEXINSHEET', -1))  # Default to -1 if not specified
    text = record.get('TEXT', '')  # Default to empty string if not specified
    show_border = record.get('SHOWBORDER', 'F')  # Default to false if not specified

    data = f"|RECORD=28|OWNERPARTID={owner_part_id}|LOCATION.X={location_x}|LOCATION.Y={location_y}|CORNER.X={corner_x}|CORNER.Y={corner_y}|AREACOLOR={area_color}|FONTID={font_id}|ALIGNMENT={alignment}|WORDWRAP={word_wrap}|COLOR={color}|ISSOLID={is_solid}|CLIPTORECT={clip_to_rect}|ISNOTACCESIBLE={is_not_accessible}|ORIENTATION={orientation}|TEXTMARGIN_FRAC={text_margin_frac}|INDEXINSHEET={index_in_sheet}|TEXT={text}|SHOWBORDER={show_border}".encode('utf-8')
    stream.write(data)
    
    stream.write(b'\n')  # Add a newline to separate records

def handle_sheet(stream, record):
    """
    Handles the Sheet record, which defines the properties for the entire schematic.
    """
    font_id_count = int(record.get('FONTIDCOUNT', 1))  # Default to 1 if not specified
    sheet_style = int(record.get('SHEETSTYLE', 0))  # Default to 0 (A4) if not specified
    system_font = int(record.get('SYSTEMFONT', 1))  # Default to 1 if not specified
    border_on = record.get('BORDERON', 'T')
    title_block_on = record.get('TITLEBLOCKON', 'T')
    snap_grid_on = record.get('SNAPGRIDON', 'T')
    visible_grid_on = record.get('VISIBLEGRIDON', 'T')
    snap_grid_size = int(record.get('SNAPGRIDSIZE', 10))  # Default to 10 if not specified
    visible_grid_size = int(record.get('VISIBLEGRIDSIZE', 10))  # Default to 10 if not specified
    custom_x = int(record.get('CUSTOMX', 0))  # Default to 0 if not specified
    custom_y = int(record.get('CUSTOMY', 0))  # Default to 0 if not specified
    use_custom_sheet = record.get('USECUSTOMSHEET', 'F')
    workspace_orientation = int(record.get('WORKSPACEORIENTATION', 0))  # Default to 0 if not specified
    custom_x_zones = int(record.get('CUSTOMXZONES', 1))  # Default to 1 if not specified
    custom_y_zones = int(record.get('CUSTOMYZONES', 1))  # Default to 1 if not specified
    custom_margin_width = int(record.get('CUSTOMMARGINWIDTH', 20))  # Default to 20 if not specified
    display_unit = int(record.get('DISPLAY_UNIT', 4))  # Default to 4 (1/100") if not specified
    reference_zones_on = record.get('REFERENCEZONESON', 'T')
    show_template_graphics = record.get('SHOWTEMPLATEGRAPHICS', 'T')
    template_filename = record.get('TEMPLATEFILENAME', '')
    area_color = int(record.get('AREACOLOR', 16777215))  # Default white if not specified

    # Assemble the data for the stream
    data = (
        f"|RECORD=31|FONTIDCOUNT={font_id_count}|SHEETSTYLE={sheet_style}|SYSTEMFONT={system_font}"
        f"|BORDERON={border_on}|TITLEBLOCKON={title_block_on}|SNAPGRIDON={snap_grid_on}"
        f"|VISIBLEGRIDON={visible_grid_on}|SNAPGRIDSIZE={snap_grid_size}|VISIBLEGRIDSIZE={visible_grid_size}"
        f"|CUSTOMX={custom_x}|CUSTOMY={custom_y}|USECUSTOMSHEET={use_custom_sheet}"
        f"|WORKSPACEORIENTATION={workspace_orientation}|CUSTOMXZONES={custom_x_zones}|CUSTOMYZONES={custom_y_zones}"
        f"|CUSTOMMARGINWIDTH={custom_margin_width}|DISPLAY_UNIT={display_unit}|REFERENCEZONESON={reference_zones_on}"
        f"|SHOWTEMPLATEGRAPHICS={show_template_graphics}|TEMPLATEFILENAME={template_filename}|AREACOLOR={area_color}"
    ).encode('utf-8')
    stream.write(data)
    
    stream.write(b'\n')  # Add a newline to separate records

def handle_sheet_name_and_file_name(stream, record):
    """
    Handles both the Sheet name and Sheet file name records, which label the schematic.
    """
    record_type = record['RECORD']  # This should be either 32 or 33
    owner_index = record.get('OWNERINDEX', 0)  # Default to 0 if not specified
    index_in_sheet = int(record.get('INDEXINSHEET', -1))  # Default to -1 if not specified
    owner_part_id = int(record.get('OWNERPARTID', -1))  # Default to -1 if not specified
    location_x = int(record.get('LOCATION.X', 0))  # Default to 0 if not specified
    location_y = int(record.get('LOCATION.Y', 0))  # Default to 0 if not specified
    color = int(record.get('COLOR', 0))  # Default to black if not specified
    font_id = int(record.get('FONTID', 1))  # Default to 1 if not specified
    text = record.get('TEXT', '')

    # Assemble the data for the stream
    data = (
        f"|RECORD={record_type}|OWNERINDEX={owner_index}|INDEXINSHEET={index_in_sheet}"
        f"|OWNERPARTID={owner_part_id}|LOCATION.X={location_x}|LOCATION.Y={location_y}"
        f"|COLOR={color}|FONTID={font_id}|TEXT={text}"
    ).encode('utf-8')
    stream.write(data)
    
    stream.write(b'\n')  # Add a newline to separate records


def write_text_frame_record(stream, record):
    try:
        loc_x = struct.pack('<I', int(record['LOCATION.X']))
        loc_y = struct.pack('<I', int(record['LOCATION.Y']))
        text = record['TEXT'].encode('utf-8')

        data = b'|RECORD=28' + b'|LOCATION.X=' + loc_x + b'|LOCATION.Y=' + loc_y + b'|TEXT=' + text + b'|'
        stream.write(data)
    except KeyError as e:
        print(f"Missing expected property for text frame: {str(e)}")

def handle_implementation_list(stream, record):
    """
    Handles the Implementation list record, listing various model types related to the schematic.
    """
    owner_index = record.get('OWNERINDEX', 0)  # Default to 0 if not specified
    
    # Assemble the data for the stream
    data = (
        f"|RECORD=44|OWNERINDEX={owner_index}"
    ).encode('utf-8')
    stream.write(data)
    stream.write(b'\n')  # Add a newline to separate records

def handle_implementation(stream, record):
    """
    Handles Implementation record for detailed model information such as footprints or 3D models.
    """
    owner_index = record.get('OWNERINDEX', 0)  # Default to 0 if not specified
    index_in_sheet = int(record.get('INDEXINSHEET', -1))  # Default to -1 if not specified
    unique_id = record.get('UNIQUEID', '')
    description = record.get('DESCRIPTION', '')
    use_component_library = record.get('USECOMPONENTLIBRARY', 'F')
    model_name = record.get('MODELNAME', '')
    model_type = record.get('MODELTYPE', '')
    data_file_count = int(record.get('DATAFILECOUNT', 0))
    model_data_file_entity0 = record.get('MODELDATAFILEENTITY0', '')
    model_data_file_kind0 = record.get('MODELDATAFILEKIND0', '')
    data_links_locked = record.get('DATALINKSLOCKED', 'F')
    database_data_links_locked = record.get('DATABASEDATALINKSLOCKED', 'F')
    integrated_model = record.get('INTEGRATEDMODEL', 'F')
    is_current = record.get('ISCURRENT', 'F')

    # Assemble the data for the stream
    data = (
        f"|RECORD=45|OWNERINDEX={owner_index}|INDEXINSHEET={index_in_sheet}|UNIQUEID={unique_id}"
        f"|DESCRIPTION={description}|USECOMPONENTLIBRARY={use_component_library}"
        f"|MODELNAME={model_name}|MODELTYPE={model_type}|DATAFILECOUNT={data_file_count}"
        f"|MODELDATAFILEENTITY0={model_data_file_entity0}|MODELDATAFILEKIND0={model_data_file_kind0}"
        f"|DATALINKSLOCKED={data_links_locked}|DATABASEDATALINKSLOCKED={database_data_links_locked}"
        f"|INTEGRATEDMODEL={integrated_model}|ISCURRENT={is_current}"
    ).encode('utf-8')
    stream.write(data)
    stream.write(b'\n')  # Add a newline to separate records

def handle_46(stream, record):
    """
    Handles specific child details of an implementation record.
    """
    owner_index = record.get('OWNERINDEX', 0)  # Default to 0 if not specified
    
    # Assemble the data for the stream
    data = (
        f"|RECORD=46|OWNERINDEX={owner_index}"
    ).encode('utf-8')
    stream.write(data)
    stream.write(b'\n')  # Add a newline to separate records

def handle_47(stream, record):
    """
    Handles the design interface and its implementations.
    """
    owner_index = record.get('OWNERINDEX', 0)  # Default to 0 if not specified
    index_in_sheet = record.get('INDEXINSHEET', 0)  # Default to 0 if not specified
    des_intf = record.get('DESINTF', '')  # Interface description, if any
    des_imp_count = record.get('DESIMPCOUNT', 1)  # Default to 1 if not specified
    des_imp0 = record.get('DESIMP0', '')  # First implementation detail

    # Assemble the data for the stream
    data = (
        f"|RECORD=47|OWNERINDEX={owner_index}|INDEXINSHEET={index_in_sheet}"
        f"|DESINTF={des_intf}|DESIMPCOUNT={des_imp_count}|DESIMP0={des_imp0}"
    ).encode('utf-8')
    stream.write(data)
    stream.write(b'\n')  # Add a newline to separate records

def handle_48(stream, record):
    """
    Handles specific implementation details as a child of a broader implementation record.
    """
    owner_index = record.get('OWNERINDEX', 0)  # Default to 0 if not specified

    # Assemble the data for the stream
    data = f"|RECORD=48|OWNERINDEX={owner_index}".encode('utf-8')
    stream.write(data)
    stream.write(b'\n')  # Add a newline to separate records

def handle_hyperlink(stream, record):
    """
    Handles hyperlink data for inclusion in the schematic file stream.
    """
    index_in_sheet = record.get('INDEXINSHEET', -1)  # Default to -1 if not specified
    owner_part_id = record.get('OWNERPARTID', -1)    # Default to -1 if not specified
    location_x = record.get('LOCATION.X', 0)         # Default location if not provided
    location_y = record.get('LOCATION.Y', 0)         # Default location if not provided
    filename = record.get('FILENAME', '').encode('utf-8')  # Ensure filename is in bytes

    # Assemble the data for the stream
    data = f"|RECORD=226|INDEXINSHEET={index_in_sheet}|OWNERPARTID={owner_part_id}|LOCATION.X={location_x}|LOCATION.Y={location_y}|FILENAME={filename}".encode('utf-8')
    stream.write(data)
    stream.write(b'\n')  # Add a newline to separate records

def handle_additional_sheet_symbol(stream, record):
    """
    Handles additional sheet symbol data, similar to the main sheet symbol.
    """
    # Here you would process and write additional details specific to a sheet symbol
    # This function should mimic the behavior of the sheet symbol with extra data

def handle_additional_generic(stream, record):
    """
    Handles generic additional record which might not have specific attributes.
    """
    # This could be used for RECORD=216 where no specific attributes are provided

def handle_additional_sheet_name_and_file_name(stream, record):
    """
    Handles additional sheet name and file name records.
    """
    owner_index_additional = record.get('OWNERINDEXADDITIONALLIST', False)
    hidden = record.get('ISHIDDEN', 'F')  # Assume not hidden if not specified
    text = record.get('TEXT', '').encode('utf-8')

    # Prepare and write the data to stream
    data = f"|RECORD=217|OWNERINDEXADDITIONALLIST={owner_index_additional}|ISHIDDEN={hidden}|TEXT={text}".encode('utf-8')
    stream.write(data)
    stream.write(b'\n')

def handle_additional_geometry(stream, record):
    """
    Handles additional geometric shapes or lines.
    """
    color = record.get('COLOR', 15187117)  # Default color
    index_in_sheet = record.get('INDEXINSHEET', -1)
    line_width = record.get('LINEWIDTH', 1)
    location_count = record.get('LOCATIONCOUNT', 2)
    owner_part_id = record.get('OWNERPARTID', -1)
    x1 = record.get('X1', 0)
    x2 = record.get('X2', 0)
    y1 = record.get('Y1', 0)
    y2 = record.get('Y2', 0)

    # Assemble the data for the stream
    data = f"|RECORD=218|COLOR={color}|INDEXINSHEET={index_in_sheet}|LINEWIDTH={line_width}|LOCATIONCOUNT={location_count}|OWNERPARTID={owner_part_id}|X1={x1}|X2={x2}|Y1={y1}|Y2={y2}".encode('utf-8')
    stream.write(data)
    stream.write(b'\n')

def append_if_present(data, record, key, format_string=None, transform=None, prefix=''):
    """
    Appends formatted data to the list if key exists in the record.

    Args:
    data (list): List of byte strings being built for the record.
    record (dict): Dictionary containing the record's data.
    key (str): Key to check in the record dictionary.
    format_string (str): struct format string to use for packing the data, if needed.
    transform (callable): Optional function to transform the value before appending.
    prefix (str): Prefix to add before the key-value pair in the data.
    """
    if key in record:
        value = record[key]
        # Apply transformation if provided
        if transform:
            value = transform(value)
        # Handle integer values that need packing
        if format_string:
            packed_value = struct.pack(format_string, int(value))
            data.append(f"|{prefix}{key}=".encode('ascii') + packed_value)
        else:
            # Directly append byte-encoded values
            data.append(f"|{prefix}{key}={value}".encode('ascii'))
        logging.debug(f"Appended {key} with value {value} to data.")



record_handlers = {
    0: handle_header,
    1: handle_component,
    2: handle_pin,
    3: handle_ieee_symbol,
    4: handle_label,
    5: handle_bezier,
    6: handle_polyline,
    7: handle_polygon,
    8: handle_ellipse,
    9: handle_piechart,
    10: handle_round_rectangle,
    11: handle_elliptical_arc,
    12: handle_arc,
    13: handle_line,
    14: handle_rectangle,
    15: handle_sheet_symbol,
    16: handle_sheet_entry,
    17: handle_power_port,
    18: handle_port,
    22: handle_no_erc,
    25: handle_net_label,
    26: handle_bus,
    27: handle_wire,
    28: handle_text_frame,
    29: handle_junction,
    30: handle_image,
    31: handle_sheet,
    32: handle_sheet_name_and_file_name,  # Assuming 32 and 33 are handled by the same function
    33: handle_sheet_name_and_file_name,
    34: handle_designator,
    37: handle_bus_entry,
    39: handle_template,
    41: handle_parameter,
    43: handle_warning_sign,
    44: handle_implementation_list,
    45: handle_implementation,
    46: handle_46,
    47: handle_47,
    48: handle_48,
    215: handle_additional_sheet_symbol,
    216: handle_additional_generic,
    217: handle_additional_sheet_name_and_file_name,
    218: handle_additional_geometry,
    226: handle_hyperlink
}

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python jsontoschdoc.py <input_json_path> <output_schdoc_path>")
    else:
        json_to_schdoc(sys.argv[1], sys.argv[2])
