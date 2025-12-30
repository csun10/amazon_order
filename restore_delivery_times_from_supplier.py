import json
import os
import glob

def restore_delivery_times():
    """Restore delivery times based on supplier information."""
    
    json_pattern = os.path.join('order_generation', 'json_template', '*.json')
    json_files = glob.glob(json_pattern)
    
    print(f"Found {len(json_files)} JSON files to process...")
    print()
    
    updated_count = 0
    
    for json_file in json_files:
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Get supplier name
            supplier_value = data.get('cells', {}).get('B3', {}).get('value', '')
            current_delivery = data.get('cells', {}).get('B14', {}).get('value', '')
            
            # Skip if already has '天' suffix
            if current_delivery and current_delivery.endswith('天'):
                continue
            
            # Determine delivery time based on supplier
            new_days = 15  # default
            
            if '印刷厂' in supplier_value:
                new_days = 7
            elif supplier_value in [
                '宁波泰丰机械有限公司',
                '阳江骏业工贸有限公司', 
                '宁波瑾秀制刷科技有限公司',
                '宁波市海曙硕丰塑料五金制品有限公司'
            ]:
                new_days = 45
            
            # Update the delivery time value
            if 'B14' in data.get('cells', {}):
                data['cells']['B14']['value'] = f"{new_days}天"
                
                # Save the updated file
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                
                print(f"✓ {os.path.basename(json_file):40} Supplier: {supplier_value[:30]:30} -> {new_days}天")
                updated_count += 1
                
        except Exception as e:
            print(f"✗ Error processing {os.path.basename(json_file)}: {e}")
    
    print()
    print(f"Completed! Updated {updated_count} files.")

if __name__ == "__main__":
    restore_delivery_times()
