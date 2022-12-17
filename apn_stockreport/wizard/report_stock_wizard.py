import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
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

    # region "All Excel Functions"
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
        product_ids, products_variants = self._get_product_attributes_variants(product_ids)
        # sales_return = self._sales_qty_returned(product_ids)
        sales_return = self._get_query_sales_return_qty_result(product_ids, start_date, end_date)

        self.fp = BytesIO()
        self.workbook = xlsxwriter.Workbook(self.fp, {'in_memory': True})
        worksheet, wbf, data_cell_formats = self._write_headers(report_name, start_date, end_date)

        # ########################### Database Operations Get Related Data ###########
        query = self._get_query(product_ids, start_date, end_date)
        self._cr.execute(query)
        result = self._cr.fetchall()
        # result = (('style code1', 20, 10, 30, 12, 11, 23, 45,), ('style code2', 30, 20, 10, 15, 14, 73, 35,),)

        # #############################  Writing Data to Excel ########################
        last_row = self._write_worksheet_data(worksheet, data_cell_formats, result, products_variants, sales_return)
        self._write_sum_of_columns(worksheet, wbf, last_row)

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

        wbf, self.workbook = self._add_workbook_format(self.workbook)

        columns_Headings = [
            ('No', 5),  # 1
            ('Style Code', 30),  # 2
            ('Description', 30),  # 3
            ('color', 30),  # 4
            ('size', 30),  # 5
            ('Opening Balance', 20),  # 6
            ('Received from Production', 30),  # 7
            ('Sales', 30),  # 8
            ('Returns', 30),  # 9
            ('Give Away/Marketing', 30),  # 10
            ('Scrap', 30),  # 11
            ('Reserved', 20),  # 12
            ('Closing Balance', 20)  # 13
        ]

        data_cell_formats = (wbf['content'], wbf['content'], wbf['content'],  # 1)No 2)Style Code 3)Description
                             wbf['content'], wbf['content'], wbf['content'],  # 4)color 5)size 6)Opening Balance
                             wbf['content'], wbf['content'], wbf['content'],  # 7)Received from Prod 8)Sales 9)Returns
                             wbf['content'], wbf['content'], wbf['content'], wbf['content'])  # 10) 11) 12) 13)

        # if 'Stock Report' not in self.workbook.:
        worksheet = self.workbook.add_worksheet(report_name)
        worksheet.merge_range('A2:M3', report_name, wbf['title_doc'])

        worksheet.write(4, 1, 'From Date', wbf['content'])
        worksheet.write(5, 1, 'To date', wbf['content'])
        worksheet.write(4, 2, start_date.strftime('%Y-%m-%d %H:%M:%S'), wbf['content_datetime'])
        worksheet.write(5, 2, end_date.strftime('%Y-%m-%d %H:%M:%S'), wbf['content_datetime'])

        row = 9
        col = 0
        for column in columns_Headings:
            column_name = column[0]
            column_width = column[1]
            worksheet.set_column(col, col, column_width)
            worksheet.write(row, col, column_name, wbf['header_orange'])
            col += 1

        return worksheet, wbf, data_cell_formats

    @staticmethod
    def _write_worksheet_data(worksheet, data_cell_formats, result, products_variants, sales_return):
        row = 10
        no = 1

        for res in result:
            # Product_id on 4 index in res. it will match in products_variants dictionary with same "product_id" key
            # and corresponding tuple will be store in product_values.This Tuple Contains following values
            # (product.id, variant, product_name_with_variant, product.default_code)
            product_values = products_variants.get(res[4])
            product_style = product_values[1]
            description = product_values[2]
            style_code = product_values[3]
            color_size = tuple(product_style.split(","))
            product_style_code = 0
            color = ''
            size = ''
            if len(color_size) >= 1:
                size = color_size[0]
            if len(color_size) >= 2:
                color = color_size[1]
            for col_number in range(13):  # 12 columns but one extra column of no in Excel so 12+1 =13
                if col_number == 0:
                    worksheet.write(row, col_number, no, data_cell_formats[0])  # Writing Serial Numbers
                elif col_number == 1:  # style code .internal reference
                    worksheet.write(row, col_number, style_code, data_cell_formats[col_number])
                elif col_number == 2:
                    worksheet.write(row, col_number, description, data_cell_formats[col_number])
                elif col_number == 3:
                    worksheet.write(row, col_number, color, data_cell_formats[col_number])
                elif col_number == 4:
                    worksheet.write(row, col_number, size, data_cell_formats[col_number])
                elif col_number == 8:
                    cell_format = data_cell_formats[col_number]
                    # return =  pos_return + sales return after picking .
                    # sales return is positive and POS return is -ve value in odoo.
                    # so multiply by -1 to make it positive.and added both qty.
                    sales_and_pos_return = sales_return.get(res[4])[1] + (-1 * res[col_number])
                    worksheet.write(row, col_number, sales_and_pos_return, cell_format)
                elif col_number == 12:  # Last Column then Formula Cell.write formula instead
                    cell_row = row + 1  # to change from (int,row,int column) to A1,B1 cell format
                    worksheet.write_formula('M%s' % cell_row,
                                            '=F%s+G%s-H%s+I%s-J%s-K%s-L%s'
                                            % (cell_row, cell_row, cell_row, cell_row, cell_row, cell_row, cell_row)
                                            )
                else:
                    # first five value in columns 0-4
                    # res = (dummy0,dummy1,dummy2,dummy3,not_used_value, value, 2, 5.0, 0.0, 4, 1.0, 0.0)
                    # res = (Prod_id,opening_balance,received_from_production,sales,returns,Give_Away_Marketing,Scrap,Reserved)
                    # Excel_columns_length =0-12
                    #
                    cell_format = data_cell_formats[col_number]
                    worksheet.write(row, col_number, res[col_number], cell_format)
            row += 1
            no += 1
        return row

    @staticmethod
    def _write_sum_of_columns(worksheet, cell_format, last_row):
        for col_number in range(12):
            if not (col_number == 0 or col_number == 1 or col_number == 2
                    or col_number == 3 or col_number == 4 or col_number == 12):
                first_cell = xl_rowcol_to_cell(10, col_number)
                last_cell = xl_rowcol_to_cell(last_row, col_number)
                worksheet.write_formula(8, col_number, '=SUM(%s:%s)' % (first_cell, last_cell),
                                        cell_format['content_number'])

    # endregion

    # region "All Queries"
    @staticmethod
    def _get_query_products(products_ids_in):
        query = f"""SELECT ID AS Prod_id FROM product_product WHERE id IN {products_ids_in} ORDER BY ID"""
        return query

    @staticmethod
    def _get_query_available(products_ids_in, start_date_string, end_date_string):
        query = f"""select prod.id AS Prod_id,
                            SUM(quant.quantity)  AS opening_balance --On Hand
                            FROM product_product prod
                            LEFT JOIN 
                            stock_quant quant  on prod.id=quant.product_id
                            LEFT JOIN 
                            stock_location loc on loc.id=quant.location_id
                            WHERE 
                            (DATE(in_date) < '{start_date_string}') --TO_DATE('{start_date_string}', 'YYYY-MM-DD' )) 
                            AND substring(
                            date_part('year',NOW())::TEXT 
                            from 3 FOR 4
                            )  
                            < '24'
                            AND 
                            loc.usage = 'internal' 
                            AND  
                            prod.id in {products_ids_in}                          
                            GROUP BY prod.id
                            ORDER BY prod.id"""
        return query

    @staticmethod
    def _get_query_reserved(products_ids_in, start_date_string, end_date_string):
        query = f"""select prod.id AS Prod_id,
                                   SUM(quant.reserved_quantity) AS Reserved --reserved
                                   FROM product_product prod
                                   LEFT JOIN 
                                   stock_quant quant  on prod.id=quant.product_id
                                   LEFT JOIN 
                                   stock_location loc on loc.id=quant.location_id
                                   WHERE 
                                   (DATE(in_date)  BETWEEN '{start_date_string}' AND '{end_date_string}') 
                                   AND substring(
                                   date_part('year',NOW())::TEXT 
                                   from 3 FOR 4
                                   )  
                                   < '24'
                                   AND 
                                   loc.usage ='internal' 
                                   AND  
                                   prod.id in {products_ids_in}                          
                                   GROUP BY prod.id
                                   ORDER BY prod.id"""
        return query
        pass

    @staticmethod
    def _get_query_scrap_qty(products_ids_in, start_date_string, end_date_string):
        # query = f"""SELECT product_id Prod_id,SUM(scrap_qty) as scrap_Qty from stock_scrap
        #             where state= 'done' AND (DATE(date_done)  BETWEEN '{start_date_string}' AND '{end_date_string}')
        #             AND substring(
        #             date_part('year',NOW())::TEXT
        #             from 3 FOR 4)
        #              < '24'
        #             AND product_id in {products_ids_in}
        #             GROUP BY product_id
        #             ORDER BY product_id
        #             """

        query = f"""select prod.id AS Prod_id,
                                   SUM(quant.quantity)  AS scrap_Qty 
                                   FROM product_product 
                                   prod
                                   LEFT JOIN 
                                   stock_quant quant  on prod.id=quant.product_id
                                   LEFT JOIN 
                                   stock_location loc on loc.id=quant.location_id
                                   WHERE 
                                   (DATE(in_date)  BETWEEN '{start_date_string}' AND '{end_date_string}') 
                                   AND substring(
                                   date_part('year',NOW())::TEXT 
                                   from 3 FOR 4
                                   )  
                                   < '24'
                                   AND 
                                   loc.usage ='inventory' 
                                   AND loc.scrap_location = TRUE
                                   AND  
                                   prod.id in {products_ids_in}                          
                                   GROUP BY prod.id
                                   ORDER BY prod.id"""
        return query

    @staticmethod
    def _get_query_sale_order_qty(products_ids_in, start_date_string, end_date_string):
        query = f"""SELECT product_id as  Prod_id , SUM(qty_invoiced) AS sales_QTY FROM sale_order_line 
                    where invoice_status = 'invoiced'
                    AND (DATE(create_date)  BETWEEN '{start_date_string}' AND '{end_date_string}') 
                    AND substring(
                    date_part('year',NOW())::TEXT 
                    from 3 FOR 4) 
                     < '24'
                    AND  product_id IN {products_ids_in}                    
                    GROUP BY product_id
                    ORDER BY product_id"""
        return query

    @staticmethod
    def _get_query_pos_order_qty(products_ids_in, start_date_string, end_date_string):
        query = f""" select product_id AS Prod_id,sum(qty) AS pos_qty  from pos_order_line
                      where qty >0
                      AND (DATE(create_date)  BETWEEN '{start_date_string}' AND '{end_date_string}') 
                      AND substring(
                      date_part('year',NOW())::TEXT 
                      from 3 FOR 4)  
                      < '24'
                      AND  product_id IN {products_ids_in}                      
                      GROUP BY product_id
                      ORDER BY product_id"""
        return query

    @staticmethod
    def _get_query_pos_order_return(products_ids_in, start_date_string, end_date_string):
        query = f"""select product_id AS Prod_id,sum(qty) AS pos_qty_return  from pos_order_line
                    where qty <0
                    AND (DATE(create_date)  BETWEEN '{start_date_string}' AND '{end_date_string}')
                    AND substring(
                    date_part('year',NOW())::TEXT
                     from 3 FOR 4)  
                    < '24' 
                    AND  product_id IN {products_ids_in}                    
                    GROUP BY product_id
                    ORDER BY product_id"""
        return query

    @staticmethod
    def _get_query_received_from_production(products_ids_in, start_date_string, end_date_string):
        query = f"""select product_id AS Prod_id,SUM(product_uom_qty) AS production_qty
                               from stock_move 
                               WHERE  product_id in {products_ids_in}
                               AND (DATE(DATE)  BETWEEN '{start_date_string}' AND '{end_date_string}')  
                               AND location_id in 
                               (select id from stock_LOCATION where Lower(NAME) LIKE Lower('%production%') AND 
                               usage ='production' and active= TRue) 
                               AND  location_dest_id in 
                               (select id from stock_LOCATION where usage = 'internal')
                               AND substring(
                               date_part('year',NOW())::TEXT
                                from 3 FOR 4)  
                               < '24'
                               GROUP BY product_id
                               ORDER BY product_id """
        return query

    @staticmethod
    def _get_query_give_away_marketing_qty(products_ids_in, start_date_string, end_date_string):
        query = f"""
                    select product_id AS Prod_id,SUM(product_uom_qty) AS marketing_qty
                    from stock_move 
                    where  product_id in {products_ids_in}
                    AND (DATE(DATE)  BETWEEN '{start_date_string}' AND '{end_date_string}')  
                    AND substring(
                    date_part('year',NOW())::TEXT 
                    from 3 FOR 4)
                      < '24'
                    AND  location_dest_id in (1765,1762)

                    GROUP BY product_id
                    ORDER BY product_id
                    """
        return query

    def _get_query_sales_return_qty_result(self, product_ids, start_date, end_date):
        products_ids_in = self._get_values_in(product_ids)
        start_date_string, end_date_string = self._get_date_string(start_date, end_date)
        sales_returned_qty = {}
        for product_id in product_ids:
            sales_returned_qty[product_id] = (product_id, 0,)

        query = f"""
                           select product_id AS Prod_id,SUM(product_uom_qty) AS sales_return
                           from stock_move 
                           where  product_id in {products_ids_in}
                           AND  reference like 'RO%' AND state ='done'
                           AND (DATE(DATE)  BETWEEN '{start_date_string}' AND '{end_date_string}')  
                           AND substring(
                           date_part('year',NOW())::TEXT 
                           from 3 FOR 4)
                             < '24'
                           GROUP BY product_id
                           ORDER BY product_id
                           """
        self._cr.execute(query)
        results = self._cr.fetchall()
        for result in results:
            product_update = {result[0]: (result[0], result[1],)}
            sales_returned_qty.update(product_update)
        return sales_returned_qty
        # return query

    def _get_query(self, product_ids, start_date, end_date):

        # purchase_locations = self._get_locations('supplier')
        # sales_Locations = self._get_locations('customer')
        # scrap_location = self._get_locations('inventory', True)
        # Inventory_Adjustment = self._get_locations('')
        products_ids_in = self._get_values_in(product_ids)
        start_date_string = start_date.strftime("%Y-%m-%d")
        end_date_string = end_date.strftime("%Y-%m-%d")
        start_date_string, end_date_string = self._get_date_string(start_date, end_date)

        query = f"""WITH 
                            cte_products AS (
                            {self._get_query_products(products_ids_in)}
                            ),

                            cte_available AS (
                            {self._get_query_available(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_scrap AS (
                            {self._get_query_scrap_qty(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_sales_order_qty AS (
                            {self._get_query_sale_order_qty(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_pos_order_qty AS (
                            {self._get_query_pos_order_qty(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_pos_qty_return AS (
                            {self._get_query_pos_order_return(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_received_from_production AS  (
                            {self._get_query_received_from_production(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_give_away_marketing_qty AS (
                            {self._get_query_give_away_marketing_qty(products_ids_in, start_date_string, end_date_string)}
                            ),

                            cte_query_reserved AS (
                            {self._get_query_reserved(products_ids_in, start_date_string, end_date_string)}
                            )


                            SELECT '0dummy' AS col0,
                            '1dummy' AS col1,
                            '2dummy' As col2,
                            '3dummy' AS col3,
                            prod.Prod_id,
                            COALESCE(available.opening_balance,0) AS opening_balance,
                            COALESCE(production.production_qty,0) AS received_from_production,
                            -- COALESCE(sales_QTY.sales_QTY,0) AS sales_qty,
                            --COALESCE(pos_qty.pos_qty,0) AS pos_qty,
                            (COALESCE(sales_QTY.sales_QTY,0)+ COALESCE(pos_qty.pos_qty,0) ) as sales,
                            COALESCE(pos_return.pos_qty_return,0) AS returns,
                            COALESCE(marketing.marketing_qty,0) AS Give_Away_Marketing,
                            COALESCE(scrap.scrap_Qty,0) AS Scrap,
                            --(COALESCE(prod.Prod_id,0)+ COALESCE(scrap_Qty,0) ) as sum,
                            COALESCE(reserved.Reserved,0) AS Reserved


                            FROM  
                            cte_products AS prod
                            LEFT JOIN  
                            cte_available AS available ON available.Prod_id = prod.Prod_id
                            LEFT JOIN 
                            cte_scrap AS scrap ON scrap.Prod_id = prod.Prod_id
                            LEFT JOIN
                            cte_sales_order_qty AS sales_QTY ON sales_QTY.Prod_id = prod.Prod_id
                            LEFT JOIN
                            cte_pos_order_qty AS pos_qty ON pos_qty.Prod_id = prod.Prod_id
                            LEFT JOIN 
                            cte_pos_qty_return AS pos_return ON pos_return.Prod_id = prod.Prod_id
                            LEFT JOIN 
                            cte_received_from_production AS production on production.Prod_id = prod.Prod_id
                            LEFT JOIN 
                            cte_give_away_marketing_qty AS marketing on marketing.Prod_id = prod.Prod_id
                            LEFT JOIN 
                            cte_query_reserved AS reserved on reserved.Prod_id = prod.Prod_id                            
                        """
        return query

    # endregion

    # region "Helper Functions"
    @staticmethod
    def _validate_data(product_ids, start_date, end_date):
        # if not product_ids:
        #     raise ValueError(_("Please Choose at least one product!"))
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

    def _get_product_attributes_variants(self, product_ids):
        if not product_ids:
            products = self.env['product.product'].search([], order='id asc')
            product_ids = [prod.id for prod in products]
        products_with_attribute = self.env['product.product'].search([('id', 'in', product_ids)], order='id asc')
        products_attributes = {}  # Store Sorted product with variant names and product names
        for product in products_with_attribute:
            variant = product.product_template_attribute_value_ids._get_combination_name()
            product_name_with_variant = variant and "%s (%s)" % (product.name, variant) or product.name
            product_info = (product.id, variant, product_name_with_variant, product.default_code)
            products_attributes[product.id] = product_info
        return product_ids, products_attributes

    @staticmethod
    def _get_values_in(values):
        string_value = str(tuple(values)).replace(',)', ')')
        return string_value

    @staticmethod
    def _get_date_string(start_date, end_date):
        start_date_string = start_date.strftime("%Y-%m-%d")
        end_date_string = end_date.strftime("%Y-%m-%d")
        return start_date_string, end_date_string

    @staticmethod
    def _add_workbook_format(workbook):

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
        wbf['content_number'].set_top()
        wbf['content_number'].set_bottom()

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

    # endregion

    # region "Extra Functions"
    @api.model
    def get_default_date_model(self):
        return pytz.UTC.localize(datetime.now()).astimezone(timezone(self.env.user.tz or 'UTC'))

    def _get_internal_transfer_locations(self):
        pass

    def _get_storable_products(self, product_ids=None):
        if not product_ids:
            product_ids = self.env['product.product'].search().id
        storable_components = self.env['product.product'].search(
            [('id', 'in', list(product_ids)), ('type', '=', 'product')])
        return storable_components

    def _get_locations(self, usage, scrap=False):
        query = """
                       select id 
                      -- ,name,complete_name,usage,scrap_location 
                       from stock_location 
                       where usage='%s' and scrap_location = %s
                   """
        self._cr.execute(query % (usage, scrap))
        result = self._cr.fetchall()

        location_ids = self.env['stock.location'].search([('usage', '=', usage), ('scrap_location', '=', scrap)])
        locations = [loc.id for loc in location_ids]
        return tuple(locations)

    def _get_available_qty(self, start_date):
        product_context = dict(request.env.context, to_date=start_date)
        product_category = request.env['product.public.category']

        # if category:
        category = product_category.browse().exists()

        attrib_list = request.httprequest.args.getlist('attrib')
        attrib_values = [[int(x) for x in v.split("-")] for v in attrib_list if v]
        attrib_set = {v[1] for v in attrib_values}

    def _sales_qty_returned(self, product_ids):
        sales_returned_qty = {}
        for product_id in product_ids:
            sales_returned_qty[product_id] = (product_id, 0,)

        lines = self.env['sale.order.line'].search([('id', 'in', product_ids)], order='id asc')
        for line in lines:
            qty = 0.0
            if line.qty_delivered_method == "stock_move":
                _, incoming_moves = line._get_outgoing_incoming_moves()
                for move in incoming_moves:
                    if move.state != "done":
                        continue
                    qty += move.product_uom._compute_quantity(
                        move.product_uom_qty,
                        line.product_uom,
                        rounding_method="HALF-UP",
                    )
            product_update = {line.product_id.id: (line.product_id.id, qty,)}
            sales_returned_qty.update(product_update)
        return sales_returned_qty
    # endregion
