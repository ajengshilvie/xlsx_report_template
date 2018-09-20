from openerp import fields, models, api
from datetime import datetime
import base64

class xlsx_report_template(models.TransientModel):
    _name = 'xlsx.report.template'

    def _get_report_title(self,worksheet,report_name,all_col,wbf,row):
        f_col = 'A'
        l_col = self._get_alphabet(all_col-1)
        return worksheet.merge_range('%s%s:%s%s'%(f_col,row,l_col,row), report_name, wbf['title_doc'])
                
    @api.multi
    def _get_converted_date(self,date):
        if not date :
            return False
        return datetime.strftime(datetime.strptime(date,'%Y-%m-%d'),'%d %B %Y')

    @api.multi
    def _get_date_and_company(self,filename,company_id=None):
        get_date = self.env['res.company'].get_default_date_model()
        date = get_date.strftime("%d-%m-%Y %H:%M:%S")
        if company_id :
            company = self.env['res.company'].browse(company_id[0])
        else :
            company = self.env.user.company_id
        company_name = company.name
        filename = filename+str(date)+'.xlsx'
        return date,company_name,filename
    
    @api.multi
    def _return_to_form(self,report,workbook,fp,filename,module,reference,class_name):
        
        workbook.close()
        out=base64.encodestring(fp.getvalue())
        report.write({'state_x':'get', 'data_x':out, 'name': filename})
        fp.close()
                
        ir_model_data = self.env['ir.model.data']
        form_res = ir_model_data.get_object_reference(module,reference)

        form_id = form_res and form_res[1] or False
        return {
            'name': 'Download XLS',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': class_name,
            'res_id':report._ids[0],
            'view_id': False,
            'views': [(form_id, 'form')],
            'type': 'ir.actions.act_window',
            'target': 'current'
        }
         
                      
    @api.multi
    def _to_date_format(self,value):
        date = datetime.strptime(value[0:19], "%Y-%m-%d %H:%M:%S")
        return date
    
    @api.multi
    def _generate_line(self,wbf,worksheet,line,no,row):
        col_line = 0
        for key,value in line.items() :
            if isinstance(value,datetime) :
                worksheet.write_datetime(row, col_line, value, wbf['content_datetime'])
            elif isinstance(value, int) :
                worksheet.write_number(row, col_line, value, wbf['content_number'])
            elif isinstance(value, float) :
                worksheet.write_number(row, col_line, value, wbf['content_float'])
            else :
                worksheet.write_string(row, col_line, value, wbf['content'])                                   
            col_line += 1
     
    def workbook_format(self, workbook=None, wbf=None):
        
        wbf['header'] = workbook.add_format({'bold': 1,'align': 'center','bg_color': '#d9d9d9','font_color': '#000000'})
        wbf['header'].set_font_size(10)
        wbf['header'].set_font_name('Arial')
        wbf['header'].set_text_wrap()

        wbf['header_no'] = workbook.add_format({'bold': 1,'align': 'center','bg_color': '#595959','font_color': '#ffffff'})
        wbf['header_no'].set_align('vcenter')
        wbf['header_no'].set_font_size(10)
        wbf['header_no'].set_font_name('Arial')
        wbf['header_no'].set_text_wrap()

        wbf['header_left'] = workbook.add_format({'bold': 1, 'align': 'left', 'font_color': '#000000'})
        wbf['header_left'].set_font_size(10)
        wbf['header_left'].set_font_name('Arial')

        wbf['footer'] = workbook.add_format({'align':'left'})
        wbf['footer'].set_font_size(10)
        wbf['footer'].set_font_name('Arial')
        
        wbf['content_datetime'] = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm:ss'})
        wbf['content_datetime'].set_font_size(10)
        wbf['content_datetime'].set_font_name('Arial')
                
        wbf['content_date'] = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        wbf['content_date'].set_font_size(10)
        wbf['content_date'].set_font_name('Arial')

        wbf['content_date_unlock'] = workbook.add_format({'num_format': 'dd/mm/yyyy','locked':0,})
        wbf['content_date_unlock'].set_font_size(10)
        wbf['content_date_unlock'].set_font_name('Arial')

        wbf['content_date_us_unlock'] = workbook.add_format({'num_format': 'mm/dd/yyyy','locked':0,})
        wbf['content_date_us_unlock'].set_font_size(10)
        wbf['content_date_us_unlock'].set_font_name('Arial')

        wbf['title_doc'] = workbook.add_format({'bold': 1,'align': 'center'})
        wbf['title_doc'].set_font_size(12)
        wbf['title_doc'].set_font_name('Arial')
        
        wbf['company'] = workbook.add_format({'bold': 1,'align': 'left'})
        wbf['company'].set_font_size(10)
        wbf['company'].set_font_name('Arial')
        
        wbf['content'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content'].set_font_size(10)
        wbf['content'].set_font_name('Arial')

        wbf['content_bold'] = workbook.add_format({'bold':1,'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_bold'].set_font_size(10)
        wbf['content_bold'].set_font_name('Arial')

        wbf['content_unlock'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter','locked':0,})
        wbf['content_unlock'].set_font_size(10)
        wbf['content_unlock'].set_font_name('Arial')

        wbf['content_center'] = workbook.add_format({'num_format': '@','align': 'center','text_wrap': True,'valign': 'vcenter',})
        wbf['content_center'].set_font_size(10)
        wbf['content_center'].set_font_name('Arial')

        wbf['content_center_unlock'] = workbook.add_format({'num_format': '@','align': 'center','text_wrap': True,'valign': 'vcenter','locked':0})
        wbf['content_center_unlock'].set_font_size(10)
        wbf['content_center_unlock'].set_font_name('Arial')

        wbf['content_float'] = workbook.add_format({'align': 'right','num_format': '#,##0.00'})
        wbf['content_float'].set_font_size(10)
        wbf['content_float'].set_font_name('Arial')

        wbf['content_float_bold'] = workbook.add_format({'bold':1,'align': 'right','num_format': '#,##0.00'})
        wbf['content_float_bold'].set_font_size(10)
        wbf['content_float_bold'].set_font_name('Arial')

        wbf['content_number'] = workbook.add_format({'num_format':'General','align': 'right'})
        wbf['content_number'].set_font_size(10)
        wbf['content_number'].set_font_name('Arial')

        wbf['content_number_bold'] = workbook.add_format({'bold': 1,'num_format':'General','align': 'right'})
        wbf['content_number_bold'].set_font_size(10)
        wbf['content_number_bold'].set_font_name('Arial')

        wbf['content_number_center'] = workbook.add_format({'num_format': 'General', 'align': 'center'})
        wbf['content_number_center'].set_font_size(10)
        wbf['content_number_center'].set_font_name('Arial')

        wbf['content_number_center_bold'] = workbook.add_format({'bold': 1,'num_format': 'General', 'align': 'center'})
        wbf['content_number_center_bold'].set_font_size(10)
        wbf['content_number_center_bold'].set_font_name('Arial')

        wbf['content_number_center_unlock'] = workbook.add_format({'num_format': 'General', 'align': 'center','locked':0})
        wbf['content_number_center_unlock'].set_font_size(10)
        wbf['content_number_center_unlock'].set_font_name('Arial')

        wbf['content_percent'] = workbook.add_format({'align': 'right','num_format': '0.00%'})
        wbf['content_percent'].set_font_size(10)
        wbf['content_percent'].set_font_name('Arial')
                
        wbf['total_float'] = workbook.add_format({'bold':1,'bg_color': '#c9c9c9','align': 'right','num_format': '#,##0.00'})
        wbf['total_float'].set_top()
        wbf['total_float'].set_font_size(10)  
        wbf['total_float'].set_font_name('Arial')    
        
        wbf['total_number'] = workbook.add_format({'align':'right','bg_color': '#c9c9c9','bold':1})
        wbf['total_number'].set_top()
        wbf['total_number'].set_font_size(10)
        wbf['total_number'].set_font_name('Arial')
        
        wbf['total'] = workbook.add_format({'bold':1,'bg_color': '#c9c9c9','align':'center'})
        wbf['total'].set_top()
        wbf['total'].set_font_size(10)
        wbf['total'].set_font_name('Arial')

        wbf['total_mid_float'] = workbook.add_format({'align': 'right', 'num_format': '#,##0.00'})
        wbf['total_mid_float'].set_font_size(10)
        wbf['total_mid_float'].set_font_name('Arial')

        wbf['total_mid_number'] = workbook.add_format({'align': 'right'})
        wbf['total_mid_number'].set_font_size(10)
        wbf['total_mid_number'].set_font_name('Arial')

        wbf['total_mid'] = workbook.add_format({'bold':1})
        wbf['total_mid'].set_font_size(10)
        wbf['total_mid'].set_font_name('Arial')

        wbf['total_mid_center'] = workbook.add_format({'bold':1,'align':'center'})
        wbf['total_mid_center'].set_font_size(10)
        wbf['total_mid_center'].set_font_name('Arial')
        
        wbf['header_detail_space'] = workbook.add_format({})
        wbf['header_detail_space'].set_top()
        wbf['header_detail_space'].set_font_size(10)
        wbf['header_detail_space'].set_font_name('Arial')
                
        wbf['header_detail'] = workbook.add_format({'bg_color': '#E0FFC2'})
        wbf['header_detail'].set_left()
        wbf['header_detail'].set_right()
        wbf['header_detail'].set_top()
        wbf['header_detail'].set_bottom()
        wbf['header_detail'].set_font_size(10)
        wbf['header_detail'].set_font_name('Arial')

        wbf['content_indent_0'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_0'].set_font_size(10)
        wbf['content_indent_0'].set_font_name('Arial')

        wbf['content_indent_bold_0'] = workbook.add_format({'bold':1,'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_bold_0'].set_font_size(10)
        wbf['content_indent_bold_0'].set_font_name('Arial')

        wbf['content_indent_1'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_1'].set_indent(1)
        wbf['content_indent_1'].set_font_size(10)
        wbf['content_indent_1'].set_font_name('Arial')

        wbf['content_indent_bold_1'] = workbook.add_format({'bold':1,'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_bold_1'].set_indent(1)
        wbf['content_indent_bold_1'].set_font_size(10)
        wbf['content_indent_bold_1'].set_font_name('Arial')

        wbf['content_indent_2'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_2'].set_indent(2)
        wbf['content_indent_2'].set_font_size(10)
        wbf['content_indent_2'].set_font_name('Arial')

        wbf['content_indent_bold_2'] = workbook.add_format({'bold':1,'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_bold_2'].set_indent(2)
        wbf['content_indent_bold_2'].set_font_size(10)
        wbf['content_indent_bold_2'].set_font_name('Arial')

        wbf['content_indent_3'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_3'].set_indent(3)
        wbf['content_indent_3'].set_font_size(10)
        wbf['content_indent_3'].set_font_name('Arial')

        wbf['content_indent_bold_3'] = workbook.add_format({'bold':1,'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_bold_3'].set_indent(3)
        wbf['content_indent_bold_3'].set_font_size(10)
        wbf['content_indent_bold_3'].set_font_name('Arial')

        wbf['content_indent_4'] = workbook.add_format({'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_4'].set_indent(4)
        wbf['content_indent_4'].set_font_size(10)
        wbf['content_indent_4'].set_font_name('Arial')

        wbf['content_indent_bold_4'] = workbook.add_format({'bold':1,'num_format': '@','text_wrap': True,'valign': 'vcenter',})
        wbf['content_indent_bold_4'].set_indent(4)
        wbf['content_indent_bold_4'].set_font_size(10)
        wbf['content_indent_bold_4'].set_font_name('Arial')
        return workbook,wbf
         
    @api.multi
    def _get_alphabet(self,key):
        alphabet = {0: 'A',
                    1: 'B',
                    2: 'C',
                    3: 'D',
                    4: 'E',
                    5: 'F',
                    6: 'G',
                    7: 'H',
                    8: 'I',
                    9: 'J',
                    10: 'K',
                    11: 'L',
                    12: 'M',
                    13: 'N',
                    14: 'O',
                    15: 'P',
                    16: 'Q',
                    17: 'R',
                    18: 'S',
                    19: 'T',
                    20: 'U',
                    21: 'V',
                    22: 'W',
                    23: 'X',
                    24: 'Y',
                    25: 'Z',
                    26: 'AA',
                    27: 'AB',
                    28: 'AC',
                    29: 'AD',
                    30: 'AE',
                    31: 'AF',
                    32: 'AG',
                    33: 'AH',
                    34: 'AI',
                    35: 'AJ',
                    36: 'AK',
                    37: 'AL',
                    38: 'AM',
                    39: 'AN',
                    40: 'AO',
                    41: 'AP',
                    42: 'AQ',
                    43: 'AR',
                    44: 'AS',
                    45: 'AT',
                    46: 'AU',
                    47: 'AV',
                    48: 'AW',
                    49: 'AX',
                    50: 'AY',
                    51: 'AZ',
                    52: 'BA',
                    53: 'BB',
                    54: 'BC',
                    55: 'BD',
                    56: 'BE',
                    57: 'BF',
                    58: 'BG',
                    59: 'BH',
                    60: 'BI',
                    61: 'BJ',
                    62: 'BK',
                    63: 'BL',
                    64: 'BM',
                    65: 'BN',
                    66: 'BO',
                    67: 'BP',
                    68: 'BQ',
                    69: 'BR',
                    70: 'BS',
                    71: 'BT',
                    72: 'BU',
                    73: 'BV',
                    74: 'BW',
                    75: 'BX',
                    76: 'BY',
                    77: 'BZ',
                    78: 'CA',
                    79: 'CB',
                    80: 'CC',
                    81: 'CD',
                    82: 'CE',
                    83: 'CF',
                    84: 'CG',
                    85: 'CH',
                    86: 'CI',
                    87: 'CJ',
                    88: 'CK',
                    89: 'CL',
                    90: 'CM',
                    91: 'CN',
                    92: 'CO',
                    93: 'CP',
                    94: 'CQ',
                    95: 'CR',
                    96: 'CS',
                    97: 'CT',
                    98: 'CU',
                    99: 'CV',
                    100: 'CW',
                    101: 'CX',
                    102: 'CY',
                    103: 'CZ',
                    104: 'DA',
                    105: 'DB',
                    106: 'DC',
                    107: 'DD',
                    108: 'DE',
                    109: 'DF',
                    110: 'DG',
                    111: 'DH',
                    112: 'DI',
                    113: 'DJ',
                    114: 'DK',
                    115: 'DL',
                    116: 'DM',
                    117: 'DN',
                    118: 'DO',
                    119: 'DP',
                    120: 'DQ',
                    121: 'DR',
                    122: 'DS',
                    123: 'DT',
                    124: 'DU',
                    125: 'DV',
                    126: 'DW',
                    127: 'DX',
                    128: 'DY',
                    129: 'DZ',
                    130: 'EA',
                    131: 'EB',
                    132: 'EC',
                    133: 'ED',
                    134: 'EE',
                    135: 'EF',
                    136: 'EG',
                    137: 'EH',
                    138: 'EI',
                    139: 'EJ',
                    140: 'EK',
                    141: 'EL',
                    142: 'EM',
                    143: 'EN',
                    144: 'EO',
                    145: 'EP',
                    146: 'EQ',
                    147: 'ER',
                    148: 'ES',
                    149: 'ET',
                    150: 'EU',
                    151: 'EV',
                    152: 'EW',
                    153: 'EX',
                    154: 'EY',
                    155: 'EZ',
                    156: 'FA',
                    157: 'FB',
                    158: 'FC',
                    159: 'FD',
                    160: 'FE',
                    161: 'FF',
                    162: 'FG',
                    163: 'FH',
                    164: 'FI',
                    165: 'FJ',
                    166: 'FK',
                    167: 'FL',
                    168: 'FM',
                    169: 'FN',
                    170: 'FO',
                    171: 'FP',
                    172: 'FQ',
                    173: 'FR',
                    174: 'FS',
                    175: 'FT',
                    176: 'FU',
                    177: 'FV',
                    178: 'FW',
                    179: 'FX',
                    180: 'FY',
                    181: 'FZ',
                    182: 'GA',
                    183: 'GB',
                    184: 'GC',
                    185: 'GD',
                    186: 'GE',
                    187: 'GF',
                    188: 'GG',
                    189: 'GH',
                    190: 'GI',
                    191: 'GJ',
                    192: 'GK',
                    193: 'GL',
                    194: 'GM',
                    195: 'GN',
                    196: 'GO',
                    197: 'GP',
                    198: 'GQ',
                    199: 'GR',
                    200: 'GS',
                    201: 'GT',
                    202: 'GU',
                    203: 'GV',
                    204: 'GW',
                    205: 'GX',
                    206: 'GY',
                    207: 'GZ',
                    208: 'HA',
                    209: 'HB',
                    210: 'HC',
                    211: 'HD',
                    212: 'HE',
                    213: 'HF',
                    214: 'HG',
                    215: 'HH',
                    216: 'HI',
                    217: 'HJ',
                    218: 'HK',
                    219: 'HL',
                    220: 'HM',
                    221: 'HN',
                    222: 'HO',
                    223: 'HP',
                    224: 'HQ',
                    225: 'HR',
                    226: 'HS',
                    227: 'HT',
                    228: 'HU',
                    229: 'HV',
                    230: 'HW',
                    231: 'HX',
                    232: 'HY',
                    233: 'HZ',
                    234: 'IA',
                    235: 'IB',
                    236: 'IC',
                    237: 'ID',
                    238: 'IE',
                    239: 'IF',
                    240: 'IG',
                    241: 'IH',
                    242: 'II',
                    243: 'IJ',
                    244: 'IK',
                    245: 'IL',
                    246: 'IM',
                    247: 'IN',
                    248: 'IO',
                    249: 'IP',
                    250: 'IQ',
                    251: 'IR',
                    252: 'IS',
                    253: 'IT',
                    254: 'IU',
                    255: 'IV',
                    256: 'IW',
                    257: 'IX',
                    258: 'IY',
                    259: 'IZ',
                    260: 'JA',
                    261: 'JB',
                    262: 'JC',
                    263: 'JD',
                    264: 'JE',
                    265: 'JF',
                    266: 'JG',
                    267: 'JH',
                    268: 'JI',
                    269: 'JJ',
                    270: 'JK',
                    271: 'JL',
                    272: 'JM',
                    273: 'JN',
                    274: 'JO',
                    275: 'JP',
                    276: 'JQ',
                    277: 'JR',
                    278: 'JS',
                    279: 'JT',
                    280: 'JU',
                    281: 'JV',
                    282: 'JW',
                    283: 'JX',
                    284: 'JY',
                    285: 'JZ',
                    286: 'KA',
                    287: 'KB',
                    288: 'KC',
                    289: 'KD',
                    290: 'KE',
                    291: 'KF',
                    292: 'KG',
                    293: 'KH',
                    294: 'KI',
                    295: 'KJ',
                    296: 'KK',
                    297: 'KL',
                    298: 'KM',
                    299: 'KN',
                    300: 'KO',
                    301: 'KP',
                    302: 'KQ',
                    303: 'KR',
                    304: 'KS',
                    305: 'KT',
                    306: 'KU',
                    307: 'KV',
                    308: 'KW',
                    309: 'KX',
                    310: 'KY',
                    311: 'KZ',
                    312: 'LA',
                    313: 'LB',
                    314: 'LC',
                    315: 'LD',
                    }
        return alphabet[key]       
