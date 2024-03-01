import pandas as pd
import yaml
from collections import OrderedDict

# A custom representer for OrderedDict that treats it as a regular dict
def dict_representer(dumper, data):
    return dumper.represent_dict(data.items())

# Add the custom representer to the Dumper
yaml.add_representer(OrderedDict, dict_representer)

# Read the Excel file
df = pd.read_excel('pipelines.xlsx')

# Prepare the YAML content
content = OrderedDict([
    ('variables', {
        'script_build': 'Scripts\\build.ps1',
    }),
    ('resources', {
        'pipelines': [],
    }),
    ('pool', {
        'vmImage': 'windows-latest',
    }),
    ('steps', []),
])

# Add each pipeline to the YAML content
for i, row in df.iterrows():
    variable_prefix = f'pipeline_{i+1}'
    content['variables'][f'{variable_prefix}'] = row['pipeline']
    content['variables'][f'{variable_prefix}Source'] = row['pipelineSource']
    content['variables'][f'{variable_prefix}JsonFile'] = row['jsonFile']
    content['resources']['pipelines'].append(OrderedDict([
        ('pipeline', f'${{ variables.{variable_prefix} }}'),
        ('source', f'${{ variables.{variable_prefix}Source }}'),
        ('trigger', True),
    ]))
    content['steps'].append(OrderedDict([
        ('task', f'${{ variables.{variable_prefix} }}'),
        ('inputs', OrderedDict([
            ('targetType', 'filePath'),
            ('filePath', '${{ variables.script_build }}'),
            ('arguments', f'-JsonFilePath "${{ variables.{variable_prefix}JsonFile }}"'),
        ])),
        ('condition', f'eq(resources.pipeline.${{variables.{variable_prefix}}}.result, \'Succeeded\')'),
    ]))

# Write the YAML file
with open('pipelines.yaml', 'w') as f:
    yaml.dump(content, f)