import math
import logging

def round_annual_reintro(number):
    return math.ceil(number)

def round_outputs(number, decimals, round_style = 'nearest'):
    '''
    Rounds outputs as specified
    rounds to nearest by default
    To round up/down set round_style to 'up' or 'down', respeciivley   
    '''
    if round_style == 'nearest':
        return round(number, decimals)
    elif round_style == up:
        return math.ceil(number)
    elif round_style == 'down':
        return math.floor(number)
    else:
        return logging.warning(f'Incorrect rounding style specified')
     