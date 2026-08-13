from openpyxl import load_workbook
import warnings
from prompt_toolkit.key_binding.bindings.named_commands import self_insert
import os

warnings.filterwarnings(
"ignore",
message="Data Validation extension is not supported and will be removed"
)

os.system("taskkill /F /IM excel.exe 1>nul 2>nul")
path = rf'C:\Users\02703821\Elanco\CH - Bestellung Monitoring\Templatka do pythona GTS\(sold-to  change  DE01)   CMD_template4.1.4.xlsm'

class licencja:

    def __init__(self,CN):
        self.CN = CN
        self.wb = load_workbook(path, keep_vba=True)
        self.ws = self.wb.worksheets[0]
        self.ws['A12'] = 'DE01'
        self.ws['B12'] = 'Change'
        self.ws['C12'] = 'Sold-to'
        self.ws['E5'] = CN

    def C33(self):

        self.ws['E23'] =  'C33'
        self.ws['E59'] = 'Yes'
        self.ws['E60'] = 'C33 - Vet Samples'
        self.ws['E61'] = 'NA'
        self.ws['E62'] = '31.12.2026'
        self.ws['E63'] = 'NA'
        self.ws['E64'] = 'CA5536030GQZ1, CA5537030GQZ1, CA5538030GQZ1, CA5539030GQZ1'
        self.ws['E65'] = '2'
        self.sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN} create C33 licence.xlsm'
        self.wb.save(self.sciezka)

    def C06(self):

        self.ws['E23'] =  'C06'
        self.ws['E59'] = 'Yes'
        self.ws['E60'] = 'C06 - Veterinary'
        self.ws['E61'] = 'NA'
        self.ws['E62'] = '30.12.9999'
        self.ws['E63'] = 'L01, L02, L03, L04, NONE'
        self.ws['E64'] = 'NA'
        self.ws['E65'] = '9,999,999'
        self.sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN} ceate C06 licence.xlsm'
        self.wb.save(self.sciezka)

    def C34(self):

        self.ws['E23'] =  'C34'
        self.ws['E59'] = 'Yes'
        self.ws['E60'] = 'C34 - Registration for Complementary Feed for Farm Animals'
        self.ws['E61'] = 'NA'
        self.ws['E62'] =  '30.12.9999'
        self.ws['E63'] = 'L10'
        self.ws['E64'] = 'NA'
        self.ws['E65'] = '9,999,999'
        self.sciezka = rf'C:\Users\02703821\OneDrive - Elanco\Desktop\robocze\{self.CN} create C34 licence.xlsm'
        self.wb.save(self.sciezka)


licencja("123").C33()