'''config functions'''
import json

def load_config(config_path='test_config.json'):
    """Load configuration from JSON file"""
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config
