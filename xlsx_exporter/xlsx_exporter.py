"""
    @package
    Output: Excel (2007-365)
    Grouped By: Value and Footprint (Package)
    Sorted By: Ref
    Fields: Item, Qty, Reference(s), Package, Value, Params, [shops from bom_exporter_config], Total Qty.

    Command line:
    python "pathToFile/xlsx_exporter.py" "%I" "%O.xlsx"
"""

import xlsxwriter
import string
import sys
from xlsx_exporter_config import *
sys.path.append(KICAD_SCRIPTING_PATH)
import kicad_netlist_reader


def from_netlist_text(text: str):
    if sys.platform.startswith('win32'):
        try:
            return text.encode('utf-8').decode('cp1252')
        except UnicodeDecodeError:
            return text
    else:
        return text


if __name__ == '__main__':
    # Generate an instance of a generic netlist, and load the netlist tree from
    # the command line option. If the file doesn't exist, execution will stop
    net = kicad_netlist_reader.netlist(sys.argv[1])

    try:
        # create a new document
        workbook = xlsxwriter.Workbook(sys.argv[2])
        workbook.set_properties({
            'title': 'Bill of Materials',
            'subject': 'Components for {}'.format(sys.argv[2].split('.')[0]),
            'author': 'KiCAD BOM Exporter',
            'company': DOCUMENT_COMPONY_NAME,
            'keywords': 'BOM, PCB, Components',
            'comments': 'Bill of Materials of {}'.format(sys.argv[2].split('.')[0]),
        })
        # create a spreadsheet
        worksheet = workbook.add_worksheet()
        worksheet.name = sys.argv[1].split('.')[0]
        worksheet.tab_color = PRIMARY_COLOR
        worksheet.zoom = 100
        # create formats
        bold = workbook.add_format(
            {
                'bold': True,
                'font': FONT_NAME,
                'font_size': FONT_SIZE,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': SECONDARY_COLOR
            })
        header = workbook.add_format(
            {
                'bold': True,
                'font': FONT_NAME,
                'font_size': FONT_SIZE,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': PRIMARY_COLOR
            })
        value = workbook.add_format(
            {
                'text_wrap': True,
                'font': FONT_NAME,
                'font_size': FONT_SIZE,
                'valign': 'vcenter'
            })
        value_centred = workbook.add_format(
            {
                'text_wrap': True,
                'font': FONT_NAME,
                'font_size': FONT_SIZE,
                'align': 'center',
                'valign': 'vcenter'
            })
        value_centred_total = workbook.add_format(
            {
                'text_wrap': True,
                'font': FONT_NAME,
                'font_size': FONT_SIZE,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': TOTAL_COLOR
            })
        value_centred_optional = workbook.add_format({
            'text_wrap': True,
            'font': FONT_NAME,
            'font_size': FONT_SIZE,
            'align': 'center',
            'valign': 'center',
            'bg_color': OPTIONAL_COLOR
        })
        value_centred_exchangeable = workbook.add_format({
            'text_wrap': True,
            'font': FONT_NAME,
            'font_size': FONT_SIZE,
            'align': 'center',
            'valign': 'center',
            'bg_color': EXCHANGEABLE_COLOR
        })

        columns: list = ['Item', 'Qty', 'Reference(s)', 'Package', 'Value', 'Params'] + list(SHOPS.keys()) + ['Total Qty.']

        index: int = 1
        for c in columns:
            worksheet.write('{}1'.format(string.ascii_uppercase[index - 1]), c, header)
            index += 1

        worksheet.write('M1', 'PCBs:', header)
        worksheet.write('M2', 1, bold)

        # let's group all the components first
        grouped = net.groupComponents(net.getInterestingComponents())

        comp_index: int = 1
        index = 0
        letter: str = 'A'
        for group in grouped:
            refs: str = ''
            comp: kicad_netlist_reader.comp
            for component in group:
                if len(refs) != 0:
                    refs += ', '
                refs += from_netlist_text(component.getRef())
                comp = component
            # Item
            worksheet.set_column('A:A', 5)
            worksheet.write('A{}'.format(comp_index + 1), comp_index, bold)
            # Qty
            worksheet.set_column('B:B', 5)
            worksheet.write('B{}'.format(comp_index + 1), len(group), value_centred)
            # Reference(s)
            worksheet.set_column('C:C', 30)
            worksheet.write('C{}'.format(comp_index + 1), refs, value)
            # Packages
            worksheet.set_column('D:D', 20)
            worksheet.write('D{}'.format(comp_index + 1), from_netlist_text(comp.getFootprint()).split(':')[1], value) # TODO:
            # Value
            worksheet.set_column('E:E', 20)
            worksheet.write('E{}'.format(comp_index + 1), from_netlist_text(comp.getValue()), value_centred)
            # Param
            param: str = ''
            voltage = comp.getField('Voltage')
            coefficient = comp.getField('Coefficient')
            if voltage != '' and coefficient != '':
                param = voltage + ', ' + coefficient
            elif voltage != '':
                param = voltage
            elif coefficient != '':
                param = coefficient
            worksheet.set_column('F:F', 15)
            worksheet.write('F{}'.format(comp_index + 1), param, value_centred)
            index = 6
            # filling shops
            for key in SHOPS.keys():
                worksheet.set_column('{}:{}'.format(string.ascii_uppercase[index], string.ascii_uppercase[index]), 15)
                worksheet.write('{}{}'.format(string.ascii_uppercase[index], comp_index + 1),
                                make_hyperlink(key, comp.getField(key)),
                                value_centred)
                index = index + 1
            # Total Qty
            worksheet.set_column('{}:{}'.format(string.ascii_uppercase[index], string.ascii_uppercase[index]), 10)
            worksheet.write_formula('{}{}'.format(string.ascii_uppercase[index], comp_index + 1),
                                    '=$M$2*B{}'.format(comp_index + 1), value_centred_total, '')
            comp_index += 1

        # polishing
        worksheet.conditional_format('E2:E{}'.format(comp_index + 1),
                                     {
                                         'type': 'text',
                                         'criteria': 'containing',
                                         'value': '(-)',
                                         'format': value_centred_optional
                                     })
        worksheet.conditional_format('E2:E{}'.format(comp_index + 1),
                                     {
                                         'type': 'text',
                                         'criteria': 'containing',
                                         'value': '(!)',
                                         'format': value_centred_exchangeable
                                     })
        worksheet.set_default_row(hide_unused_rows=True)
        workbook.close()
    except IOError:
        print(__file__, ':', 'Can\'t open output file for writing:' + sys.argv[2], sys.stderr)
