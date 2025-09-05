import json

def load_config(config_path='test_config.json'):
    """Load configuration from JSON file"""
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config


config = load_config()


csv_data = {'Scenario_number': 1, 'ab': 2}
print(type(csv_data))
list_csv = list(csv_data.values())
print(csv_data)        
print(list_csv)   