from odoo import api, models, fields
from cStringIO import StringIO
import xlsxwriter
from collections import OrderedDict
import time

class ReportIncomingShipment(models.TransientModel):
    _name           = "asn.report.incoming.shipment"
    _description    = "Report Incoming Shipments"

    name                = fields.Char(Filename = "Name")
    product_ids         = fields.Many2many('product.product','report_stock_product_rel','report_id','product_id',string='Product')
    start_date          = fields.Date('Start Date')
    end_date            = fields.Date('End Date',default=time.strftime('%Y-%m-%d'))
    categ_id            = fields.Many2one('product.category','Category Product')
    state_x             = fields.Selection([('choose', 'choose'), ('get', 'get')], default = 'choose')
    data_x              = fields.Binary('File', readonly = True)
    wbf = {}

    @api.multi
    def _set_query_where(self):
        query_where = " WHERE b.usage='internal' "
        if self.product_ids:
            product_ids = self.product_ids.ids
            query_where += ' and e.id in %s' % str(tuple(product_ids)).replace(',)',')')
        if self.start_date:
            query_where += " and a.in_date >= '%s 00:00:00'" % str(self.start_date)
        if self.end_date:
            query_where += " and a.in_date <= '%s 23:59:59'" % str(self.end_date)
        if self.categ_id :
            categ_id = self.categ_id.ids
            query_where += " and pc.id in %s" % str(tuple(categ_id)).replace(',)',')')
        return query_where

    @api.multi
    def _excecute_query(self, query_where):
        query = """
        select 	    
            x.name as nama_product, 
            b.complete_name as location,   
            a.in_date as in_date,  
            d.name as batch_code,  
            a.qty as qty,  
            a.cost as cost,  
            h.name as p_categ_name,  
            g.name as variant 
        From  
            stock_quant a 
            LEFT JOIN stock_location b ON b.id = a.location_id 
            LEFT JOIN stock_production_lot d ON d.id = a.lot_id 
            LEFT JOIN product_product e ON e.id = a.product_id 
            LEFT JOIN product_template x ON x.id = e.product_tmpl_id 
            LEFT JOIN product_category pc ON pc.id = x.categ_id
            LEFT JOIN product_attribute_value_product_product_rel f ON f.product_product_id = a.product_id 
            LEFT JOIN product_attribute_value g ON g.id = f.product_attribute_value_id 
            LEFT JOIN product_category h ON h.id = x.categ_id 
        %s  
        order by x.name asc
        """%query_where
        self._cr.execute(query)
        ress = self._cr.fetchall()
        return ress

    @api.multi
    def get_header_title(self,wbf):
        header_title = OrderedDict()
        header_title['No'] = [8, wbf['content_number_center']]
        header_title['Product'] = [30, wbf['content']]
        header_title['Stock Location'] = [35, wbf['content']]
        header_title['In Date'] = [18, wbf['content_date']]
        header_title['Batch Code'] = [20, wbf['content']]
        header_title['Qty'] = [10, wbf['content_number']]
        header_title['Cost'] = [15, wbf['content_float']]
        header_title['Category'] = [15, wbf['content']]
        header_title['Variant'] = [15, wbf['content']]
        return header_title

    @api.multi
    def excel_report(self):
        ### Set Template ######################
        template = self.env['xlsx.report.template']
        fp = StringIO()
        workbook = xlsxwriter.Workbook(fp)
        workbook, wbf = template.workbook_format(workbook, self.wbf)
        
        #set font and size
        workbook.formats[0].font_name = 'Arial'
        workbook.formats[0].font_size = 10
        
        #set filename for report
        filename = 'Report Incoming Shipment.xlsx'
        
        #set worksheet
        worksheet = workbook.add_worksheet('Incoming')

        #set company for title of report
        template._get_report_title(worksheet, self.env.user.company_id.name.upper(), 9, wbf, 1)

        #get and set query
        query_where = self._set_query_where()
        ress = self._excecute_query(query_where)

        #set header row
        header_title = self.get_header_title(wbf)
        row = 3
        header_row = row
        col = 0
        worksheet.set_column(col, col, len(header_title))
        for key, value in header_title.items():
            worksheet.set_column(col, col, value[0], value[1])
            worksheet.write_string(row, col, key, wbf['header_no'])
            col += 1
        last_col = col - 1

        #set data row after header + 1
        row = 4
        no = 1
        for res in ress:
            col_detail = 1
            worksheet.write(row, 0, no)
            for data in range(len(header_title)-1):
                worksheet.write(row, col_detail, res[data] if res[data] else '')
                col_detail += 1
                data += 1
            no += 1
            row += 1

        #set autofilter and freeze panes
        worksheet.autofilter(header_row, 2, row, last_col)
        worksheet.freeze_panes(header_row + 1, 3)

        # Module Alias
        module_name = 'asn_report_incoming_shipment'
        reference = 'view_asn_incoming_shipment'
        class_name = self.__class__.__name__

        return template._return_to_form(self, workbook, fp, filename, module_name, reference, class_name)

