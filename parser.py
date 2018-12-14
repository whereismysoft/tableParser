# before run:
#     create env (pip3)
#     install package docx
# if env created run comand: source tutorial-env/bin/activate (Linux) || tutorial-env\Scripts\activate.bat (Windows)
# then run file

from docx.api import Document

document = Document('text.docx')

outputDoc = Document()

table = document.tables[0]

data = []

keys = None
for i, row in enumerate(table.rows):
    text = list(cell.text for cell in row.cells)

    values = []
    # Establish the mapping based on the first row
    # headers; these will become the keys of our dictionary
    if i == 0:
        # keys = tuple(text)
        keys = tuple(('id','building', 'names'))
        continue
    values.append(i)
    values.append(text[1])
    values.append(text[2])

    for n, el in enumerate(values):
        if (n == 1):
            values[n] = el.replace('\n', '')
        if (n == 2):
            values[n] = el.split('\n')
        
    # Construct a dictionary for this row, mapping
    # keys to values for this row
    row_data = dict(zip(keys, values))
    data.append(row_data)

outputDoc.add_paragraph(str(data))

outputDoc.save('demo.docx')
