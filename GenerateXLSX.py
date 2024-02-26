import yaml
import pandas as pd

# Load the YAML file
with open('pipelines.yaml', 'r') as f:
    data = yaml.safe_load(f)

# Extract the pipeline data
pipelines_data = []
for i in range(1, len(data['resources']['pipelines']) + 1):
    pipeline = data['variables'][f'pipeline_{i}']
    pipelineSource = data['variables'][f'pipeline_{i}Source']
    jsonFile = data['variables'][f'pipeline_{i}JsonFile']
    pipelines_data.append([pipeline, pipelineSource, jsonFile])

# Convert the data to a DataFrame
df = pd.DataFrame(pipelines_data, columns=['pipeline', 'pipelineSource', 'jsonFile'])

# Write the DataFrame to an Excel file
df.to_excel('pipelines.xlsx', index=False)
