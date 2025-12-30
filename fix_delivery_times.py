import json
import os
import glob
import re

def add_tian_to_delivery_time(json_file_path):
    """Add '天' suffix to delivery time in JSON template files."""
    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Check if B14 cell exists and has a delivery time value
        if 'B14' in data.get('cells', {}):
            current_value = str(data['cells']['B14'].get('value', ''))
            
            # Skip if empty or already has '天'
            if not current_value or current_value.endswith('天'):
                return False
            
            # Check if it's a date format (contains '年' or '月' or '日')
            if '年' in current_value or '月' in current_value or '日' in current_value:
                # Extract year and calculate days from today
                import re
                from datetime import datetime
                
                # Try to parse the date
                try:
                    # Extract date components
                    match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', current_value)
                    if match:
                        year, month, day = map(int, match.groups())
                        target_date = datetime(year, month, day)
                        today = datetime.now()
                        days_diff = (target_date - today).days
                        
                        if days_diff > 0:
                            new_value = f"{days_diff}天"
                            data['cells']['B14']['value'] = new_value
                            
                            with open(json_file_path, 'w', encoding='utf-8') as f:
                                json.dump(data, f, ensure_ascii=False, indent=2)
                            
                            print(f"Updated {os.path.basename(json_file_path)}: {current_value} -> {new_value}")
                            return True
                except:
                    pass
                
                # If date parsing failed, skip this file
                print(f"Skipped {os.path.basename(json_file_path)}: Date format detected but couldn't parse")
                return False
            
            # Extract numbers if there are any non-numeric characters
            numbers = re.findall(r'\d+', current_value)
            if numbers:
                new_value = numbers[0] + '天'
                
                data['cells']['B14']['value'] = new_value
                
                # Save the updated file
                with open(json_file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                
                print(f"Updated {os.path.basename(json_file_path)}: {current_value} -> {new_value}")
                return True
            else:
                # If no numbers found, skip this file
                return False
        else:
            return False
            
    except Exception as e:
        print(f"Error processing {json_file_path}: {e}")
        return False

def main():
    # Get all JSON files in the json_template directory
    json_pattern = os.path.join(os.path.dirname(__file__), 'order_generation', 'json_template', '*.json')
    json_files = glob.glob(json_pattern)
    
    print(f"Found {len(json_files)} JSON files to process...")
    print()
    
    updated_count = 0
    skipped_count = 0
    for json_file in json_files:
        result = add_tian_to_delivery_time(json_file)
        if result:
            updated_count += 1
        else:
            skipped_count += 1
    
    print()
    print(f"Completed! Updated {updated_count} files, skipped {skipped_count} files.")

if __name__ == "__main__":
    main()
