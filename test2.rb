require 'rubygems'
require 'write_xlsx'
require 'csv'
require 'pry'

workbook = WriteXLSX.new('tracker.xlsx')

# Add sheets with names
programmes = workbook.add_worksheet('Programmes (h1)')
projects = workbook.add_worksheet('Projects (h2)')
grants = workbook.add_worksheet('Grants (h3) + spend data')
lookups = workbook.add_worksheet('Lookups')

# Defined names
workbook.define_name('Country_names', 'Lookups!$B$2:$B$252')
workbook.define_name('Country_codes', 'Lookups!$A$2:$A$252')

# Set active worksheet on Excel open
programmes.activate

# Set up some cell formats
header_bold = workbook.add_format(bold: 1, border: 1)
header_bold_grey = workbook.add_format(bold: 1, bg_color: 23, border: 1) # grey
output = workbook.add_format(bg_color: 12, border: 1) # blue
read_only = workbook.add_format(bg_color: 22, border: 1) # light grey
example_text = workbook.add_format(color: 22)

# Set up some headers
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

# Populate Lookups
countries_csv = CSV.new(File.open('countries.csv').read)
codes = []
names = []
countries_csv.read.each do |row|
  codes << row[0]
  names << row[1]
end

countries = [codes, names]
lookups.write(0, 0, countries)

# Write headers
programmes.write_row(0, 0, programmes_header, header_bold_grey)
projects.write(0, 0, projects_header, header_bold)
grants.write(0, 0, grants_header, header_bold)

# Set colours
grants.set_column(3, 0, 10, output)
grants.set_column(0, 0, 10, read_only)
projects.set_column(0, 0, 10, read_only)
programmes.set_column(0, 2, 10, read_only)

# Add list validation
# TODO: how can we make sure validation is applied to all new rows (beyond D50)?
projects.data_validation(
  'D2:D50',
  validate: 'list',
  value: "=Country_names"
)

# Add date validation
date_format = workbook.add_format(num_format: 'dd-mm-yyyy')
projects.set_column(6, 0, 10, date_format)
projects.set_column(7, 0, 10, date_format)
projects.write(1, 6, 'dd-mm-yy', example_text)
projects.write(1, 7, 'dd-mm-yy', example_text)

workbook.close