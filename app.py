from numbers_parser import Document
from target import target_filename
doc = Document(target_filename)
sheets = doc.sheets

# gets the labels of the sheets
def get_labels(first_row):
    labels = []
    for label in first_row:
        if label.value != None:
            labels.append(label.value.strip())
    return labels

# creates one entry per row
def create_entry_objects(rows):
    objects = []
    # Get labels from first row
    labels = get_labels(rows[0])
    content_rows = rows[slice(1, len(rows))]
    for row in content_rows:
        obj = {}
        for index, label in enumerate(labels):
            obj[label] = row[index].value
        objects.append(obj)
    
    return objects

# parses each sheet
def parse_sheets(sheets):
    all_sheets = []
    for sheet in sheets:
        tables = sheet.tables
        rows = tables[0].rows()
        all_sheets.append(create_entry_objects(rows))
    return all_sheets

# finds duplicates while parsing each song
def get_duplicates(prased_sheets):
    duplicate_entries = []
    matched_entries = {}
    for sheet in prased_sheets:
        for entry in sheet:
            if entry['Song'] in matched_entries and entry['Song'] != 'Break' and entry['Song'] != None:
                duplicate_entries.append(entry)
            else:
                matched_entries[entry['Song']] = entry

    return duplicate_entries

all_sheets = parse_sheets(sheets)

duplicates = get_duplicates(all_sheets)

if len(duplicates):
    print('Duplicate entries found:', duplicates)
else:
    print('No duplicates found!')