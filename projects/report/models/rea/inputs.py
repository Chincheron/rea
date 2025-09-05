from dataclasses import dataclass, field
import json

def load_config(config_path='config.json'):
    """Load configuration from JSON file"""
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config

config = load_config()
default_values = config['excel']['input_values_default']
print(default_values)

@dataclass
class REAScenarioInputs:
    '''
    Defines dataclass for REA Inputs
    Based on input_cells and input_values_default from the config file
    '''
    number_killed: int = field(default=default_values['number_killed'])
    start_year_analysis: int = field(default=default_values['start_year_analysis'])
    start_year_reproduction: int = field(default=default_values['start_year_reproduction'])
    base_year: int = field(default=default_values['base_year'])
    max_age: int = field(default=default_values['max_age'])
    discount_factor: float = field(default=default_values['discount_factor'])
    base_year: int = field(default=default_values['base_year'])
    no_reintroduction_years: int = field(default=default_values['no_reintroduction_years'])
    start_year_reintroduction: int = field(default=default_values['start_year_reintroduction'])
    annual_reintroduction: int = field(default=default_values['annual_reintroduction'])


