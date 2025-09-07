import inputs
import json
import sys

#load config
with open('config.json', 'r') as f:
        config = json.load(f)


# rea_inputs = inputs.REAScenarioInputs()
rea_inputs = inputs.REAScenarioInputs.create_from_config('config.json')

print(type(rea_inputs))
print(rea_inputs)