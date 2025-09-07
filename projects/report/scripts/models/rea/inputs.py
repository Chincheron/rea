from dataclasses import dataclass, field, asdict
import json
import pandas as pd

def load_config(config_path='config.json'):
    """Load configuration from JSON file"""
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config

config = load_config()
default_values = config['excel']['input_values_default']
print(config)

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
    no_reintroduction_years: int = field(default=default_values['no_reintroduction_years'])
    start_year_reintroduction: int = field(default=default_values['start_year_reintroduction'])
    annual_reintroduction: int = field(default=default_values['annual_reintroduction'])

    @classmethod
    def create_from_config(cls, config_path, **kwargs):
        '''Create class object from configfile'''
        config_path = 'config.json'
        config = load_config(config_path)
        default_values = config['excel']['input_values_default'] #extract dict of default values from config file
        
        field_values = {}
        for field_name in cls.__dataclass_fields__:
            if field_name in kwargs: 
                field_values[field_name] = kwargs[field_name]
            else:
                field_values[field_name] = default_values[field_name]
                print(False)
        return cls(**field_values)


    @classmethod
    def create_from_row(cls, row):
        """
        Create a REAScenarioInputs object:
        - start with defaults from config
        - override any attributes that exist in the row
        """
        obj = cls()  # start with all defaults
        for field_name in obj.__dataclass_fields__:
            if hasattr(row, field_name) and pd.notna(getattr(row, field_name)):
                setattr(obj, field_name, getattr(row, field_name))
        return obj

    def to_dict(self):
        return asdict(self)


# test = REAScenarioInputs()
# test = test.to_dict()
# print(test)
# print(type(test))


