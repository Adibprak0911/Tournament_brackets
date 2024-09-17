import pandas as pd
from itertools import combinations

def read_excel_and_generate_pairs(file_path):
    excel_file = pd.ExcelFile(file_path)
    sheet_values_by_sheet = {}

    for sheet in excel_file.sheet_names:
        sheet_data = pd.read_excel(file_path, sheet_name=sheet, header=None)
        col_index = 1
        start_row = 2
        column_values = {}
        sheet_values = {}
        
        while col_index < len(sheet_data.columns):
            values = []
            row_index = start_row
            while row_index < len(sheet_data) and pd.notna(sheet_data.iloc[row_index, col_index]):
                values.append(sheet_data.iloc[row_index, col_index])
                row_index += 1
            if values:
                column_values[col_index] = values
            col_index += 1
        for values in column_values.values():
            pool_key = values[0][0]
            if pool_key not in sheet_values:
                sheet_values[pool_key] = []
            sheet_values[pool_key].extend(list(combinations(values, 2)))
        
        sheet_values_by_sheet[sheet] = sheet_values
        # print(sheet_values_by_sheet)
                
    return sheet_values_by_sheet

def ordered_pairings(sheet_values):
    ordered_list = []
    max_len = max(len(pairs) for pairs in sheet_values.values())
    
    for i in range(max_len):
        for pool, pairings in sheet_values.items():
            if i < len(pairings):
                ordered_list.append(pairings[i])
            if len(pairings) - 1 - i > i:
                ordered_list.append(pairings[len(pairings) - 1 - i])
    seen = set()
    unique_list = [item for item in ordered_list if not (item in seen or seen.add(item))]
    print(unique_list)
    return unique_list

def check_back_to_back(team, prev_matches, current_matches):
    """Check if a team has back-to-back matches across any court."""
    for match in prev_matches:
        if team in match:
            return True
    for match in current_matches:
        if team in match:
            return True
    return False

def fill_sheet_with_ordered_list(sheet_data, ordered_list, match_index, col_index):
    row_index = 10
    available_courts = 0
    court_col_start = col_index
    
    # Check how many courts are available in the current set of columns
    while col_index < len(sheet_data.columns) and pd.notna(sheet_data.iloc[row_index-1, col_index]):
        available_courts += 1
        col_index += 1
    
    if available_courts == 0:
        return sheet_data, match_index
    
    previous_row_matches = []
    
    for row in range(10, len(sheet_data)):
        current_row_matches = []
        
        for court in range(court_col_start, court_col_start + available_courts):
            if match_index < len(ordered_list):
                match = ordered_list[match_index]
                
                # Check for back-to-back matches for both teams
                team1, team2 = match
                if check_back_to_back(team1, previous_row_matches, current_row_matches) or \
                   check_back_to_back(team2, previous_row_matches, current_row_matches):
                    # Swap the match with a later one that doesn't cause a back-to-back issue
                    for swap_index in range(match_index + 1, len(ordered_list)):
                        swap_match = ordered_list[swap_index]
                        swap_team1, swap_team2 = swap_match
                        if not (check_back_to_back(swap_team1, previous_row_matches, current_row_matches) or
                                check_back_to_back(swap_team2, previous_row_matches, current_row_matches)):
                            # Swap matches
                            ordered_list[match_index], ordered_list[swap_index] = swap_match, match
                            match = swap_match
                            break
                
                # Fill the match in the current court and row
                sheet_data.iloc[row, court] = f"{match[0]} vs {match[1]}"
                current_row_matches.append(match)
                match_index += 1
            else:
                # All matches filled, break the inner loop
                break
        
        previous_row_matches = current_row_matches
        
        if match_index >= len(ordered_list):
            break

    return sheet_data, match_index

def process_and_save_fixtures(file_path):
    sheet_values_by_sheet = read_excel_and_generate_pairs(file_path)
    excel_file = pd.ExcelFile(file_path)
    output_files = []

    for sheet_name, sheet_values in sheet_values_by_sheet.items():
        ordered_list = ordered_pairings(sheet_values)
        sheet_data = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        
        # Start filling from column 2, row 10
        match_index = 0
        col_index = 2
        
        # Fill the first set of courts (col 2, row 10)
        updated_sheet_data, match_index = fill_sheet_with_ordered_list(sheet_data, ordered_list, match_index, col_index)

        # Move to next set of columns and fill the remaining matches
        while match_index < len(ordered_list):
            col_index += 4  # Move to the next set of columns (6, 10, 14, etc.)
            
            # Check if there's enough space to move further
            if col_index >= len(sheet_data.columns):
                break
            
            updated_sheet_data, match_index = fill_sheet_with_ordered_list(updated_sheet_data, ordered_list, match_index, col_index)

        output_file_path = f"{sheet_name}_updated.xlsx"
        updated_sheet_data.to_excel(output_file_path, index=False, header=False)
        output_files.append(output_file_path)
    
    return output_files

# Example usage
file_path = 'Scores.xlsx'
process_and_save_fixtures(file_path)
