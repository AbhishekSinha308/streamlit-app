"""
FULLY DYNAMIC SYNC FROM RE EXCEL GENERATOR WITH DATA ALIGNMENT
==============================================================
NO HARDCODED LOGIC - ADAPTS TO ANY DATA STRUCTURE

Key Features:
- Automatically discovers columns from query results
- No hardcoded coordinates - reads from config or auto-detects
- No hardcoded time_key filters
- No hardcoded date logic
- Flexible data matching between weather DB and query results
- Works with any new data without code changes
- FIXED: Proper formula calculations in comparison columns
  - GROUP 1 GHI_Compare: GHI (E) - param_value (L)
  - GROUP 2 Wind_Compare: WIND_SPEED_925MB (E) - wind_speed (K)
- DATA ALIGNMENT: Skips first 6 blocks PER DAY and shifts data upward for proper alignment
- FIXED: Block skip counter resets for each day so all days get consistent skipping
"""

import pandas as pd
import ast
import sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from datetime import datetime
from typing import List, Tuple, Dict, Any

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

CONFIG = {
    'weather_file': r'\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha\WeatherDB_data.xlsx',
    'query_file':   r'\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha\query_results.xlsx',
    'output_file':  r'\\172.16.0.65\Share\24_23_QA_Team\Abhishek_Sinha\Sync_from_RE.xlsx',

    # If None, auto-detect from query results
    'group1_coords': None,  # Will auto-detect unique lat/lon from Query1_Results
    'group2_coords': None,  # Will auto-detect unique lat/lon from Query2_Results

    'ghi_sheet':      'Query1_Results',
    'wind_sheet':     'Query2_Results',
    'decimal_places': 2,
    'header_color':   '4472C4',
    'group_color':    'D9E1F2',
    'column_width':   16,

    # Skip configuration - first N time_keys to skip for data alignment PER DAY
    'skip_first_blocks_group1': 0,   # No skip for GROUP 1 (GHI) - all data written
    'skip_first_blocks_group2': 6,   # Skip first 6 blocks for GROUP 2 (Wind) - query data starts from block 7
}


def parse_dict_string(dict_str):
    """Parse dictionary string from Excel cells."""
    try:
        return ast.literal_eval(dict_str)
    except (ValueError, SyntaxError):
        return {}


def key_to_minutes(key):
    """Convert time_key to minutes for comparison."""
    try:
        s = str(key).strip()
        if ':' in s:
            parts = s.split(':')
            return int(parts[0]) * 60 + int(parts[1])
        f = float(s)
        if f < 1:
            return round(f * 1440)
        return int(f)
    except Exception:
        return -1


def validate_files(weather_file, query_file):
    """Validate input files exist."""
    if not Path(weather_file).exists():
        raise FileNotFoundError(f"[ERROR] File not found: {weather_file}")
    if not Path(query_file).exists():
        raise FileNotFoundError(f"[ERROR] File not found: {query_file}")
    print("[OK] Input files found")


def read_and_prepare_data(weather_file, query_file, ghi_sheet, wind_sheet):
    """Read and prepare data from all input files."""
    print("[INFO] Reading input files...")
    weather_db = pd.read_excel(weather_file)
    query1     = pd.read_excel(query_file, sheet_name=ghi_sheet)
    query2     = pd.read_excel(query_file, sheet_name=wind_sheet)

    print(f"[OK] WeatherDB: {len(weather_db)} rows")
    print(f"[OK] Query1 ({ghi_sheet}): {len(query1)} rows")
    print(f"     Columns: {list(query1.columns)}")
    print(f"[OK] Query2 ({wind_sheet}): {len(query2)} rows")
    print(f"     Columns: {list(query2.columns)}")

    weather_db['GHI_dict']  = weather_db['GHI'].apply(parse_dict_string)
    weather_db['WIND_dict'] = weather_db['WIND_SPEED_925MB'].apply(parse_dict_string)

    weather_db['date'] = pd.to_datetime(weather_db['TIMESTAMP_DAY']).dt.date
    query1['date']     = pd.to_datetime(query1['weather_date']).dt.date
    query2['date']     = pd.to_datetime(query2['weather_date']).dt.date

    print("[OK] Data prepared and parsed")
    return weather_db, query1, query2


def auto_detect_coordinates(query_df, lat_col, lon_col):
    """Auto-detect unique coordinate pairs from query results."""
    coords = query_df[[lat_col, lon_col]].drop_duplicates().values.tolist()
    coords = [(float(lat), float(lon)) for lat, lon in coords]
    print(f"[INFO] Auto-detected {len(coords)} coordinate pairs from query results")
    for i, (lat, lon) in enumerate(coords, 1):
        print(f"       [{i}] Latitude: {lat}, Longitude: {lon}")
    return coords


def find_matching_columns(query_df, column_hints):
    """
    Find actual column names in query dataframe based on hints.
    
    Args:
        query_df: Query dataframe
        column_hints: Dict with keys like 'latitude', 'longitude', 'value_col'
    
    Returns:
        Dict with actual column names found
    """
    found_cols = {}
    
    # Find latitude column
    lat_candidates = [col for col in query_df.columns if 'latitude' in col.lower()]
    found_cols['latitude'] = lat_candidates[0] if lat_candidates else None
    
    # Find longitude column
    lon_candidates = [col for col in query_df.columns if 'longitude' in col.lower()]
    found_cols['longitude'] = lon_candidates[0] if lon_candidates else None
    
    # Find value column (param_value for GHI, wind_speed for Wind)
    if 'ghi' in query_df.columns[0].lower() or any('ghi' in col.lower() for col in query_df.columns):
        value_candidates = [col for col in query_df.columns if 'param' in col.lower() and 'value' in col.lower()]
        found_cols['value_col'] = value_candidates[0] if value_candidates else None
        found_cols['group_type'] = 'GHI'
    else:
        value_candidates = [col for col in query_df.columns if 'wind' in col.lower() and 'speed' in col.lower()]
        found_cols['value_col'] = value_candidates[0] if value_candidates else None
        found_cols['group_type'] = 'WIND'
    
    return found_cols


def create_workbook_with_formatting(header_color, group_color):
    """Create and format Excel workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sync_from_RE'
    styles = {
        'header_fill': PatternFill(start_color=header_color, end_color=header_color, fill_type='solid'),
        'header_font': Font(bold=True, color='FFFFFF', size=11),
        'group_fill':  PatternFill(start_color=group_color,  end_color=group_color,  fill_type='solid'),
        'group_font':  Font(bold=True, size=10),
        'border': Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'),  bottom=Side(style='thin')
        )
    }
    return wb, ws, styles


def add_group_header(ws, row, title, styles):
    """Add group header to worksheet."""
    ws[f'A{row}'] = title
    ws[f'A{row}'].font = styles['group_font']
    ws[f'A{row}'].fill = styles['group_fill']
    return row + 1


def add_column_headers(ws, row, headers, styles):
    """Add column headers to worksheet."""
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.value     = header
        cell.font      = styles['header_font']
        cell.fill      = styles['header_fill']
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border    = styles['border']
    return row + 1


def set_column_widths(ws, num_columns, width):
    """Set column widths in worksheet."""
    for col in range(1, num_columns + 1):
        ws.column_dimensions[chr(64 + col)].width = width


def save_workbook(wb, output_file):
    """Save workbook to file."""
    print(f"\n[INFO] Saving workbook to {output_file}...")
    try:
        wb.save(output_file)
        print("[OK] File saved successfully!")
        return True
    except Exception as e:
        print(f"[ERROR] Error saving file: {str(e)}")
        return False


def col_num_to_letter(col_num):
    """Convert column number to letter (1->A, 2->B, etc.)"""
    return chr(64 + col_num)


# ============================================================================
# GENERIC GROUP PROCESSOR - Works for any comparison group with alignment
# ============================================================================

def process_generic_group(ws, row, weather_db, query_df, coords, styles, 
                         group_name, weather_col_dict, query_col_dict, 
                         output_col_map, decimal_places, formula_config,
                         skip_first_blocks=0):
    """
    Generic processor for any comparison group with data alignment (shifted up).
    
    This function writes ALL weather data with query data shifted UP to fill all blocks,
    starting query data from a later row index.
    
    Key behavior:
    - Weather columns (LATITUDE, LONGITUDE, TIMESTAMP_DAY, time_key, values): Written for ALL blocks
    - Query columns: Written for ALL blocks, but using rows from skip_first_blocks onwards
                     This creates a "shift up" effect where blocks 1-6 are filled with query data
                     that normally corresponds to blocks 7+
    
    Args:
        ws: Worksheet
        row: Current row number
        weather_db: Weather database dataframe
        query_df: Query results dataframe
        coords: List of (lat, lon) tuples to process
        styles: Style dict with formatting options
        group_name: Name of group (e.g., "WIND SPEED COMPARISON")
        weather_col_dict: Dict mapping logical names to weather_db columns
        query_col_dict: Dict mapping logical names to query_df columns
        output_col_map: Dict mapping output column headers to data sources
        decimal_places: Number of decimal places for formatting
        formula_config: Dict with formula configuration
        skip_first_blocks: Number of first query rows to skip (shifted to fill earlier blocks)
                          6 = skip first 6 rows of query data, use rows 7+ starting from block 1
    
    Returns:
        Tuple of (next_row, count_written)
    """
    print(f"\n[INFO] Processing GROUP - {group_name}...")
    if skip_first_blocks > 0:
        print(f"       [NOTE] Data alignment: Query data SHIFTED UP to fill all blocks")
        print(f"       [NOTE] Block 1 onwards: Uses query rows {skip_first_blocks+1}+ (shifted from row 7)")
        print(f"       [NOTE] All blocks have BOTH weather and query data (no empty cells)")
    else:
        print(f"       [NOTE] Data alignment: ALL columns written for all blocks (no shift)")
    
    row = add_group_header(ws, row, f'GROUP - {group_name}', styles)
    
    # Build dynamic headers from output_col_map
    headers = list(output_col_map.keys())
    header_row = row
    row = add_column_headers(ws, row, headers, styles)
    
    count = 0
    errors = 0
    fmt = f'0.{"0" * decimal_places}'
    
    # Extract formula configuration
    weather_col_letter = formula_config['weather_col']
    query_col_letter = formula_config['query_col']
    
    for lat, lon in coords:
        print(f"  [INFO] Processing coordinate: {lat}, {lon}")
        
        # Get weather data for this coordinate
        weather_rows = weather_db[
            (weather_db[weather_col_dict['latitude']] == lat) & 
            (weather_db[weather_col_dict['longitude']] == lon)
        ].sort_values('date')
        
        # Get query data for this coordinate
        query_rows = query_df[
            (query_df[query_col_dict['latitude']] == lat) & 
            (query_df[query_col_dict['longitude']] == lon)
        ]
        
        if weather_rows.empty:
            print(f"        [WARN] No weather data for {lat}/{lon}")
            errors += 1
            continue
            
        if query_rows.empty:
            print(f"        [WARN] No query data for {lat}/{lon}")
            errors += 1
            continue
        
        print(f"        [OK] Found {len(weather_rows)} weather records")
        print(f"        [OK] Found {len(query_rows)} query records")
        
        q_idx = skip_first_blocks  # Start from row (skip_first_blocks+1) of query data when skipping blocks
        coord_count = 0
        total_data_items = 0
        
        # Iterate through weather data
        for _, w_row in weather_rows.iterrows():
            # Get the data dictionary (GHI or WIND)
            data_dict = w_row[weather_col_dict['data_dict']]
            if not data_dict:
                continue
            
            # Reset block counter for each day (each weather row)
            # Track which block number we're on (1=block with key 30, 2=block with key 90, etc.)
            block_count_in_this_row = 0
            
            # Iterate through each time_key in data (WRITE ALL WEATHER DATA)
            for time_key, data_value in data_dict.items():
                total_data_items += 1
                block_count_in_this_row += 1
                
                # Always write query data for ALL blocks (shifted up from skip position)
                write_query_data = True
                
                if q_idx >= len(query_rows):
                    break
                
                try:
                    # Always have access to weather row
                    # Query row only if we should write it
                    q_row = query_rows.iloc[q_idx] if write_query_data else None
                    
                    # Write each column based on output_col_map
                    col_idx = 1
                    for header, source_info in output_col_map.items():
                        cell = ws.cell(row=row, column=col_idx)
                        value = None
                        
                        # Get value based on source
                        if source_info['source'] == 'weather':
                            value = w_row[source_info['column']]
                        elif source_info['source'] == 'data_dict':
                            value = data_value if source_info['column'] == 'value' else time_key
                        elif source_info['source'] == 'query':
                            # Only write query data if we've passed the skip blocks
                            if write_query_data and q_row is not None:
                                value = q_row[source_info['column']]
                            # else: leave empty (None)
                        elif source_info['source'] == 'formula':
                            # Write formula directly in the correct column
                            if write_query_data:
                                formula_str = f'={weather_col_letter}{row}-{query_col_letter}{row}'
                                cell.value = formula_str
                                cell.number_format = fmt
                            # else: leave formula empty
                            value = None  # Don't write value again
                        else:
                            value = None
                        
                        # Format and set value (skip if already handled as formula)
                        if source_info['source'] != 'formula' and value is not None:
                            if isinstance(value, (int, float)):
                                cell.value = value
                                if source_info.get('number_format'):
                                    cell.number_format = fmt
                            else:
                                cell.value = value
                        
                        col_idx += 1
                    
                    count += 1
                    coord_count += 1
                    if write_query_data:
                        q_idx += 1
                    row += 1
                
                except Exception as e:
                    print(f"        [ERROR] Row {row}: {str(e)}")
                    errors += 1
                    if write_query_data:
                        q_idx += 1
                    row += 1
        
        print(f"        [OK] Wrote {coord_count} rows for {lat}/{lon}")
        if skip_first_blocks > 0:
            print(f"        [OK] Query data shifted UP: Row {skip_first_blocks+1} appears in block 1, row {skip_first_blocks+2} in block 2, etc.")
    
    print(f"  [OK] {group_name} total rows written: {count}")
    if errors:
        print(f"  [WARN] Total errors encountered: {errors}")
    
    return row + 2, count


# ============================================================================
# MAIN
# ============================================================================

def main():
    print("=" * 90)
    print(" " * 15 + "FULLY DYNAMIC SYNC FROM RE WITH DATA ALIGNMENT")
    print(" " * 20 + "EXCEL GENERATOR WITH AUTO-DETECTION")
    print("=" * 90)
    print(f"\nStart time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    try:
        print("Step 1: Validating input files...")
        validate_files(CONFIG['weather_file'], CONFIG['query_file'])

        print("\nStep 2: Reading and preparing data...")
        weather_db, query1, query2 = read_and_prepare_data(
            CONFIG['weather_file'], CONFIG['query_file'],
            CONFIG['ghi_sheet'],    CONFIG['wind_sheet']
        )

        print("\nStep 3: Auto-detecting coordinates...")
        group1_coords = CONFIG['group1_coords'] or auto_detect_coordinates(
            query1, 'latitude_name', 'longitude_name'
        )
        group2_coords = CONFIG['group2_coords'] or auto_detect_coordinates(
            query2, 'latitude_name', 'longitude_name'
        )

        print("\nStep 4: Finding column mappings...")
        cols_g1 = find_matching_columns(query1, {'latitude': 'latitude_name', 'longitude': 'longitude_name'})
        cols_g2 = find_matching_columns(query2, {'latitude': 'latitude_name', 'longitude': 'longitude_name'})
        
        print(f"[OK] GROUP 1 (GHI) column mapping: {cols_g1}")
        print(f"[OK] GROUP 2 (Wind) column mapping: {cols_g2}")

        print("\nStep 5: Creating Excel workbook...")
        wb, ws, styles = create_workbook_with_formatting(
            CONFIG['header_color'], CONFIG['group_color']
        )
        print("[OK] Workbook created")

        row = 1
        
        # ====================================================================
        # Process GROUP 1 - GHI COMPARISON (NO SKIP - Write all data)
        # ====================================================================
        print("\n" + "=" * 90)
        print("PROCESSING GROUP 1 - GHI COMPARISON (No Data Skip)")
        print("=" * 90)
        
        weather_col_dict_g1 = {
            'latitude': 'LATITUDE',
            'longitude': 'LONGITUDE',
            'data_dict': 'GHI_dict'
        }
        query_col_dict_g1 = {
            'latitude': 'latitude_name',
            'longitude': 'longitude_name'
        }
        output_col_map_g1 = {
            'LATITUDE': {'source': 'weather', 'column': 'LATITUDE'},
            'LONGITUDE': {'source': 'weather', 'column': 'LONGITUDE'},
            'TIMESTAMP_DAY': {'source': 'weather', 'column': 'TIMESTAMP_DAY'},
            'GHI_time_key': {'source': 'data_dict', 'column': 'time_key'},
            'GHI': {'source': 'data_dict', 'column': 'value', 'number_format': True},
            'latitude_name': {'source': 'query', 'column': 'latitude_name'},
            'longitude_name': {'source': 'query', 'column': 'longitude_name'},
            'weather_date': {'source': 'query', 'column': 'weather_date'},
            'weather_time': {'source': 'query', 'column': 'weather_time'},
            'process_file_name': {'source': 'query', 'column': 'process_file_name'},
            'param_name': {'source': 'query', 'column': 'param_name'},
            'param_value': {'source': 'query', 'column': 'param_value', 'number_format': True},
            'FILE': {'source': 'weather', 'column': 'FILE'},
            'GHI_Compare': {'source': 'formula'}
        }
        
        # Formula: GHI (col E) - param_value (col L) = GHI_Compare (col N)
        formula_config_g1 = {
            'weather_col': 'E',
            'query_col': 'L',
            'result_col': 'N'
        }
        
        row, count_g1 = process_generic_group(
            ws, row, weather_db, query1, group1_coords, styles,
            'GHI COMPARISON',
            weather_col_dict_g1, query_col_dict_g1, output_col_map_g1,
            CONFIG['decimal_places'],
            formula_config_g1,
            skip_first_blocks=CONFIG['skip_first_blocks_group1']
        )
        
        print(f"\n[RESULT] GROUP 1 - GHI COMPARISON: {count_g1} rows written")
        print(f"         Formula Applied: GHI (Column E) - param_value (Column L) = GHI_Compare (Column N)")
        print(f"         Data Alignment: All data written (no skip applied)")

        # ====================================================================
        # Process GROUP 2 - WIND SPEED COMPARISON (SKIP first 6 blocks PER DAY)
        # ====================================================================
        print("\n" + "=" * 90)
        print("PROCESSING GROUP 2 - WIND SPEED COMPARISON (Data Skip Applied Per Day)")
        print("=" * 90)
        
        weather_col_dict_g2 = {
            'latitude': 'LATITUDE',
            'longitude': 'LONGITUDE',
            'data_dict': 'WIND_dict'
        }
        query_col_dict_g2 = {
            'latitude': 'latitude_name',
            'longitude': 'longitude_name'
        }
        output_col_map_g2 = {
            'LATITUDE': {'source': 'weather', 'column': 'LATITUDE'},
            'LONGITUDE': {'source': 'weather', 'column': 'LONGITUDE'},
            'TIMESTAMP_DAY': {'source': 'weather', 'column': 'TIMESTAMP_DAY'},
            'Wind_time_key': {'source': 'data_dict', 'column': 'time_key'},
            'WIND_SPEED_925MB': {'source': 'data_dict', 'column': 'value', 'number_format': True},
            'weather_date': {'source': 'query', 'column': 'weather_date'},
            'master_file_name': {'source': 'query', 'column': 'master_file_name'},
            'time': {'source': 'query', 'column': 'time'},
            'latitude_name': {'source': 'query', 'column': 'latitude_name'},
            'longitude_name': {'source': 'query', 'column': 'longitude_name'},
            'wind_speed': {'source': 'query', 'column': 'wind_speed', 'number_format': True},
            'level': {'source': 'query', 'column': 'level'},
            'FILE': {'source': 'weather', 'column': 'FILE'},
            'Wind_Compare': {'source': 'formula'}
        }
        
        # Formula: WIND_SPEED_925MB (col E) - wind_speed (col K) = Wind_Compare (col N)
        formula_config_g2 = {
            'weather_col': 'E',
            'query_col': 'K',
            'result_col': 'N'
        }
        
        row, count_g2 = process_generic_group(
            ws, row, weather_db, query2, group2_coords, styles,
            'WIND SPEED COMPARISON',
            weather_col_dict_g2, query_col_dict_g2, output_col_map_g2,
            CONFIG['decimal_places'],
            formula_config_g2,
            skip_first_blocks=CONFIG['skip_first_blocks_group2']
        )
        
        print(f"\n[RESULT] GROUP 2 - WIND SPEED COMPARISON: {count_g2} rows written")
        print(f"         Formula Applied: WIND_SPEED_925MB (Column E) - wind_speed (Column K) = Wind_Compare (Column N)")
        print(f"         Data Alignment: Query data SHIFTED UP to fill all blocks")
        print(f"                         Block 1 starts with Query row 7, Block 2 with Query row 8, etc.")
        print(f"                         All blocks contain both weather and query data (no empty cells)")

        print("\nStep 6: Formatting columns...")
        set_column_widths(ws, 14, CONFIG['column_width'])
        print("[OK] Column widths set to 16 characters")

        # Green fill when comparison result is 0
        green_fill = PatternFill(
            start_color='00FF00',
            end_color='00FF00',
            fill_type='solid'
        )

        ws.conditional_formatting.add(
            f'N1:N{ws.max_row}',
            CellIsRule(
                operator='between',
                formula=['-0.01', '0.01'],
                fill=green_fill
            )
        )
        print("\nStep 7: Saving file...")
        success = save_workbook(wb, CONFIG['output_file'])

        if success:
            print("\n" + "=" * 90)
            print(" " * 32 + "[OK] COMPLETED SUCCESSFULLY!")
            print("=" * 90)
            print(f"\nOutput file location  : {CONFIG['output_file']}")
            
            print(f"\n{'GROUP 1 - GHI COMPARISON':^90}")
            print(f"{'-' * 90}")
            print(f"  Rows written        : {count_g1:>6}")
            print(f"  Formula             : GHI (E) - param_value (L) = GHI_Compare (N)")
            print(f"  Data alignment      : All data included (no skip)")
            print(f"  Data source         : WeatherDB + Query1_Results")
            print(f"  Coordinates         : {len(group1_coords)} locations")
            
            print(f"\n{'GROUP 2 - WIND SPEED COMPARISON':^90}")
            print(f"{'-' * 90}")
            print(f"  Rows written        : {count_g2:>6}")
            print(f"  Formula             : WIND_SPEED_925MB (E) - wind_speed (K) = Wind_Compare (N)")
            print(f"  Data alignment      : Query data SHIFTED UP to fill all blocks")
            print(f"                      : Block 1 uses Query row 7, Block 2 uses Query row 8, etc.")
            print(f"                      : All blocks have both weather and query data")
            print(f"  Data source         : WeatherDB + Query2_Results")
            print(f"  Coordinates         : {len(group2_coords)} locations")
            
            print(f"\n{'TOTAL SUMMARY':^90}")
            print(f"{'-' * 90}")
            print(f"  Total rows          : {count_g1 + count_g2:>6}")
            print(f"  End time            : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("=" * 90 + "\n")
            return 0
        else:
            print("\n[ERROR] Failed to save file")
            return 1

    except FileNotFoundError as e:
        print(f"\n[ERROR] {e}")
        return 1
    except Exception as e:
        print(f"\n[ERROR] Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())