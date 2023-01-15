import docx
import pandas as pd

# Open the Word document
doc = docx.Document('test.docx')

# Define the text you're looking for
search_text = 'test'

# Initialize an empty list to store the data
data = []

# Iterate through each table in the document
for table in doc.tables:
    # Iterate through each row in the table
    for row in table.rows:
        # Initialize an empty list to store the cell data
        row_data = []
        # Iterate through each cell in the row
        for cell in row.cells:
            # Append the cell's text to the row_data list
            row_data.append(cell.text)
        # Append the row_data to the data list
        data.append(row_data)

# Create a pandas dataframe from the data
df = pd.DataFrame(data)
df.columns = ['Column1','Column2']

# Find the row where Column1 is equal to search_text
matching_row = df.loc[df['Column1'] == search_text]

# Print the corresponding text in Column2
print(matching_row['Column2'].values[0])
