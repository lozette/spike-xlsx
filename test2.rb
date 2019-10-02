require 'rubygems'
require 'write_xlsx'

workbook = WriteXLSX.new('tracker.xlsx')

programmes = workbook.add_worksheet('Programmes (h1)')
projects = workbook.add_worksheet('Projects (h2)')
grants = workbook.add_worksheet('Grants (h3) + spend data')
lookups = workbook.add_worksheet('Lookups')
programmes.activate

header_bold = workbook.add_format(bold: 1)
header_bold_grey = workbook.add_format(bold: 1, bg_color: 23)
output = workbook.add_format(bg_color: 12)
read_only = workbook.add_format(bg_color: 22)

programmes_header = ['Unique BEIS ID', 'Programme Title', 'Budget']
projects_header = [
    'Programme Title',
    'Project Title',
    'Project Description',
    'Country / Region',
    'Delivery Partner Unique ID',
    'Status',
    'Open Date',
    'Close Date',
    'Channel of Delivery Code',
    'Sector Code',
    'UK Named Contact',
    'Aid Type',
    'Climate Change Adaption',
    'Climate Change Mitigation',
    'Desertification',
    'Biodiversity',
    'Gender'
]
grants_header = [
    'Programme Title',
    'Project Title',
    'Grant Title',
    'Q3 2019/20 Actuals',
    'Q4 2019/20 Forecast',
    'Q1 2020/21 Forecast',
    'Q2 2020/21 Forecast',
    'Q3 2020/21 Forecast',
    'Q4 2020/21 Forecast',
    'Q1 2021/22 Forecast',
    'Q2 2021/22 Forecast',
    'Q3 2021/22 Forecast',
    'Q4 2021/22 Forecast'
]

programmes.write_row(0, 0, programmes_header, header_bold_grey)
projects.write(0, 0, projects_header, header_bold)
grants.write(0, 0, grants_header, header_bold)
grants.set_column(3, 0, 10, output)
grants.set_column(0, 0, 10, read_only)
projects.set_column(0, 0, 10, read_only)
programmes.set_column(0, 2, 10, read_only)

workbook.close