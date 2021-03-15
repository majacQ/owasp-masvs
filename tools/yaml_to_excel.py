import yaml
from openpyxl import load_workbook

MASVS_TITLES = {
    'V1': 'Architecture, Design and Threat Modeling Requirements',
    'V2': 'Data Storage and Privacy Requirements',
    'V3': 'Cryptography Requirements',
    'V4': 'Authentication and Session Management Requirements',
    'V5': 'Network Communication Requirements',
    'V6': 'Platform Interaction Requirements',
    'V7': 'Code Quality and Build Setting Requirements',
    'V8': 'Resilience Requirements',
}

def get_hyperlink(url):
    title = url.split('#')[1].replace('-',' ').capitalize() 
    return f'=HYPERLINK("{url}", "{title}")'

CHECKMARK = "✓"

masvs_dict = yaml.load(open('masvs_full.yaml'))

wb = load_workbook(filename='checklist.xlsx')
table = wb['Security Requirements - Android']

row=4
col_id=2
col_mstg_id=3
col_text=4
col_l1=5
col_l2=6
col_link=8

for mstg_id, req in masvs_dict.items():
    req_id = req['id'].split('.') 
    category = req_id[0]
    subindex = req_id[1]

    if subindex == '1':
        category_id = f"V{category}"
        category_title = MASVS_TITLES[category_id]
        table.cell(row=row,column=col_id).value = category_id
        table.cell(row=row,column=col_id+2).value = category_title
        row = row+1
    
    l1 = CHECKMARK if req['L1'] else ""
    l2 = CHECKMARK if req['L2'] else ""
    r = CHECKMARK if req['R'] else ""

    table.cell(row=row,column=col_id).value = req['id']
    table.cell(row=row,column=col_mstg_id).value = mstg_id
    table.cell(row=row,column=col_text).value = req['text']
    table.cell(row=row,column=col_l1).value = l1
    table.cell(row=row,column=col_l2).value = l2
    if req.get('links'):
        table.cell(row=row,column=col_link).value = get_hyperlink(req['links'][0])
    row = row+1

wb.save(filename='checklist_out.xlsx')