from openpyxl import load_workbook
import warnings

warnings.filterwarnings(
"ignore",
message="Data Validation extension is not supported and will be removed"
)

def C06(CN):
    template = rf'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'
    wb = load_workbook(template, keep_vba=True)
    ws= wb['Sheet1']

    ws['A12'] = 'DE01'
    ws['B12'] = 'Change'
    ws['C12']= 'Sold-to'

    ws['E5'] = CN
    ws['E23'] =  'C06'
    ws['E59'] = 'Yes'
    ws['E60'] = 'C06 - Veterinary'
    ws['E61'] = 'NA'
    ws['E62'] = '30.12.9999'
    ws['E63'] = 'L01, L02, L03, L04, NONE'
    ws['E64'] = 'NA'
    ws['E65'] = '9,999,999'


    sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C06 licence.xlsm'

    wb.save(sciezka)


def C33(CN):
    template = rf'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'
    wb = load_workbook(template,keep_vba=True)
    ws= wb['Sheet1']


    ws['A12'] = 'DE01'
    ws['B12'] = 'Change'
    ws['C12']= 'Sold-to'

    ws['E5'] = CN
    ws['E23'] =  'C33'
    ws['E59'] = 'Yes'
    ws['E60'] = 'C33 - Vet Samples'
    ws['E61'] = 'NA'
    ws['E62'] =  '30.12.9999'
    ws['E63'] = 'NA'
    ws['E64'] = 'CA5536030GQZ1, CA5537030GQZ1, CA5538030GQZ1, CA5539030GQZ1'
    ws['E65'] = '2'


    sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C33 licence.xlsm'

    wb.save(sciezka)


def C34(CN):
    template = rf'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'
    wb = load_workbook(template, keep_vba=True)
    ws= wb['Sheet1']

    ws['A12'] = 'DE01'
    ws['B12'] = 'Change'
    ws['C12']= 'Sold-to'

    ws['E5'] = CN
    ws['E23'] =  'C34'
    ws['E59'] = 'Yes'
    ws['E60'] = 'C33 - Vet Samples'
    ws['E61'] = 'NA'
    ws['E62'] =  '30.12.9999'
    ws['E63'] = 'NA'
    ws['E64'] = 'CA5536030GQZ1, CA5537030GQZ1, CA5538030GQZ1, CA5539030GQZ1'
    ws['E65'] = '2'


    sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{CN} create C33 licence.xlsm'

    wb.save(sciezka)
    wb.close(sciezka)

C06('2505')