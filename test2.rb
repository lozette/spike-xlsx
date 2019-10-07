require 'rubygems'
require 'write_xlsx'
require 'csv'
require 'pry'

workbook = WriteXLSX.new('tracker.xlsx')

# Add sheets with names
programmes = workbook.add_worksheet('Programmes (h1)')
projects   = workbook.add_worksheet('Projects (h2)')
grants     = workbook.add_worksheet('Grants (h3) + spend data')
lookups    = workbook.add_worksheet('Lookups')

# Defined names
workbook.define_name('Country_names', 'Lookups!$B$1:$B$252')
workbook.define_name('Country_codes', 'Lookups!$A$1:$A$252')
workbook.define_name('Statuses', 'Lookups!$C$1:$C$8')

# Set active worksheet on Excel open
programmes.activate

# Set up some custom colours
# So the first param here (index) is because you're actually overwriting
# the existing colour index (8..64) with your own colours
# So weird.
dark_grey     = workbook.set_custom_color(10, 174, 170, 170)
grey          = workbook.set_custom_color(11, 208, 206, 206)
light_grey    = workbook.set_custom_color(12, 217, 217, 217)
lightest_grey = workbook.set_custom_color(13, 242, 242, 242)
dark_blue     = workbook.set_custom_color(14, 0, 112, 192)
light_blue    = workbook.set_custom_color(15, 0, 176, 240)
purple        = workbook.set_custom_color(16, 112, 48, 160)
green         = workbook.set_custom_color(17, 169, 208, 142)
amber         = workbook.set_custom_color(18, 244, 176, 132)
red           = workbook.set_custom_color(19, 252, 104, 110)
blue          = workbook.set_custom_color(20, 68, 194, 196)

# Set up some cell formats
header_bold      = workbook.add_format(bold: 1, border: 1)
header_bold_grey = workbook.add_format(bold: 1, bg_color: dark_grey, border: 1)
output           = workbook.add_format(bg_color: light_blue, border: 1)
read_only        = workbook.add_format(bg_color: light_grey, border: 1)
example_text     = workbook.add_format(color: dark_grey)

status_grey                = workbook.add_format(bg_color: grey)
status_green               = workbook.add_format(bg_color: green)
status_red                 = workbook.add_format(bg_color: red)
status_amber               = workbook.add_format(bg_color: amber)
status_blue                = workbook.add_format(bg_color: blue)
status_no_longer_happening = workbook.add_format(bg_color: lightest_grey)
status_delivery            = workbook.add_format(bg_color: dark_grey)

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
lookups_header = [
    'Countries',
    '',
    'Statuses',
    '',
    'Pillars',
    '',
    'Match type',
    '',
    'ODA Considerations',
    '',
    'GCRF Challenge Areas',
    '',
    'Aid Types',
    '',
    'Channel of Delivery Codes',
    '',
    'Sector Codes'
]

# Populate Lookups
countries_csv = CSV.new(File.open('countries.csv').read, headers: true)
codes = []
names = []
countries_csv.read.each do |row|
  codes << row[0]
  names << row[1]
end

countries = [codes, names]
lookups.write(1, 0, countries)

# Write headers
programmes.write_row(0, 0, programmes_header, header_bold_grey)
projects.write(0, 0, projects_header, header_bold)
grants.write(0, 0, grants_header, header_bold)
lookups.write(0, 0, lookups_header, header_bold)

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

projects.data_validation(
  'F2:F50',
  validate: 'list',
  value: '=Statuses'
)

# Add boolean validation
projects.data_validation(
    'M1:M50',
    validate: 'list',
    value: [0, 1]
)

# Add status lookups
lookups.write('C2', 'GREY', status_grey)
lookups.write('C3', 'GREEN', status_green)
lookups.write('C4', 'AMBER', status_amber)
lookups.write('C5', 'RED', status_red)
lookups.write('C6', 'BLUE', status_blue)
lookups.write('C7', 'No Longer Happening', status_no_longer_happening)
lookups.write('C8', 'Delivery', status_delivery)

# Add date validation
date_format = workbook.add_format(num_format: 'dd-mm-yyyy')
projects.set_column(6, 0, 10, date_format)
projects.set_column(7, 0, 10, date_format)
projects.write(1, 6, 'dd-mm-yy', example_text)
projects.write(1, 7, 'dd-mm-yy', example_text)

workbook.close