from dataclasses import dataclass, field, asdict
import json
import pandas as pd

def load_config(config_path='config.json'):
    """Load configuration from JSON file"""
    with open(config_path, 'r') as f:
        config = json.load(f)
    return config

@dataclass
class REAScenarioInputs:
    '''
    Defines dataclass for REA Inputs
    Recommend using a config file with option create_from_config method to specify inputs
    '''
    number_killed: int 
    start_year_analysis: int 
    start_year_reproduction: int 
    discount_start_year: int 
    maximum_age: int 
    discount_factor: float 
    no_reintroduction_years: int 
    start_year_reintroduction: int 
    annual_reintroduction: int 

    @classmethod
    def create_from_config(cls, config_path, **kwargs):
        '''Create class object from configfile'''
        config = load_config(config_path)
        default_values = config['excel']['input_values_default'] #extract dict of default values from config file
        
        field_values = {}
        for field_name in cls.__dataclass_fields__:
            if field_name in kwargs: 
                field_values[field_name] = kwargs[field_name]
            else:
                field_values[field_name] = default_values.get(field_name)
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
    
    
    def update_from_row(self, row):
        """
        Update an existing REAScenarioInputs object:
        - override any attributes that exist in the row
        """
        for field_name in self.__dataclass_fields__:
            if hasattr(row, field_name) and pd.notna(getattr(row, field_name)):
                setattr(self, field_name, getattr(row, field_name))
        return self

    def to_dict(self):
        return asdict(self)


