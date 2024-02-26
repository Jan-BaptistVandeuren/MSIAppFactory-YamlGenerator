import pandas as pd
import yaml

# Read the Excel file
df = pd.read_excel('pipelines.xlsx')

# Prepare the YAML content
content = {
    'variables': {
        'script_build': 'Scripts\\build.ps1',
    },
    'resources': {
        'pipelines': [],
    },
    'pool': {
        'vmImage': 'windows-latest',
    },
    'steps': [],
}

# Add each pipeline to the YAML content
for i, row in df.iterrows():
    variable_prefix = f'pipeline_{i+1}'
    content['variables'][f'{variable_prefix}'] = row['pipeline']
    content['variables'][f'{variable_prefix}Source'] = row['pipelineSource']
    content['variables'][f'{variable_prefix}JsonFile'] = row['jsonFile']
    content['resources']['pipelines'].append({
        'pipeline': f'${{ variables.{variable_prefix} }}',
        'source': f'${{ variables.{variable_prefix}Source }}',
        'trigger': True,
    })
    content['steps'].append({
        'task': f'${{ variables.{variable_prefix} }}',
        'inputs': {
            'targetType': 'filePath',
            'filePath': '${{ variables.script_build }}',
            'arguments': f'-JsonFilePath "${{ variables.{variable_prefix}JsonFile }}"',
        },
        'condition': f'eq(resources.pipeline.${{variables.{variable_prefix}}}.result, \'Succeeded\')',
    })

# Write the YAML file
with open('pipelines.yaml', 'w') as f:
    yaml.dump(content, f)