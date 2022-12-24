# -*- coding: utf-8 -*-
"""
## Turning a dataset into a word docx table
## Daniel Girvitz 
## December 24th 2022

"""
import pandas as pd
import aspose.words as aw

# read text file into pandas DataFrame.
df = pd.read_csv("pancreatic.csv", sep=",", header=None)
print(df)

# specific to "pancreatic.csv"
del df[0]
print(df)

# Create a new Word document.
doc = aw.Document()

# Create document builder.
builder = aw.DocumentBuilder(doc)

# Start the table.
table = builder.start_table()

# Set alignment and font settings for entire table.
builder.paragraph_format.alignment = aw.ParagraphAlignment.LEFT
builder.font.size = 9
builder.font.name = "Arial"
builder.cell_format.vertical_alignment = aw.tables.CellVerticalAlignment.CENTER
builder.cell_format.width = 100.0
builder.row_format.height = aw.HeightRule.AUTO

# Set alignment and font settings for header.
builder.font.bold = True

# Create header.
for j in range(0,df.shape[1]):
    # Insert cell.
    builder.insert_cell()
    builder.write(df.iloc[0,j])
builder.end_row()

# Set alignment and font settings for rest of table.
builder.font.bold = False

# Create table.
for i in range(1,df.shape[0]):
    for j in range(0,df.shape[1]):
        builder.insert_cell()
        builder.write(df.iloc[i,j])
    builder.end_row()

# End table.
builder.end_table()

# Save the document.
doc.save("table_formatted.docx")