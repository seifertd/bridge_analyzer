#!/usr/bin/env ruby
require 'google/apis/sheets_v4'
require 'googleauth'
require 'csv'
require 'date'
require 'fileutils'

# Usage check
if ARGV.size < 1 || ARGV.size > 2
  puts "Usage: ruby update_google_sheet.rb <INPUTFILE> [PARTNER_NAME]"
  puts "  INPUTFILE: Path to a .BWS or .pbn file"
  puts "  PARTNER_NAME: (optional) Partner's first name - overrides BWS if provided"
  exit 1
end

INPUT_FILE = ARGV[0]
PARTNER_OVERRIDE = ARGV[1]

# Config
eval(File.read('./env.rb'))

# Step 1: Run the CSV generator to get partner name and CSV data
puts "Step 1: Generating CSV from Input file..."
script = if INPUT_FILE.downcase.end_with?("bws")
           "export_to_sheet.rb"
         else
           "pbn_to_results.rb"
         end
cmd = if PARTNER_OVERRIDE
  "ruby #{script} \"#{INPUT_FILE}\" \"#{PARTNER_OVERRIDE}\""
else
  "ruby #{script} \"#{INPUT_FILE}\""
end

csv_data = `#{cmd} 2>&1`
unless $?.success?
  puts "Error running CSV generator:"
  puts csv_data
  exit 1
end

puts csv_data

# Convert to CSV
csv_data = CSV.parse(csv_data)

session_date = nil
session_date_str = nil
tab_name = nil
# Step 2: Extract session date from Input file
puts "\nStep 2: Extracting session date from Input file..."
if INPUT_FILE.downcase.end_with?('bws')
  date_csv = `#{MDBEXPORT} -B "#{INPUT_FILE}" ReceivedData 2>&1`
  unless $?.success?
    puts "Error extracting date:"
    puts date_csv
    exit 1
  end

  # Parse first row to get date
  data = CSV.parse(date_csv, headers: true)
  first_row = data.first
  unless first_row && first_row['DateLog']
    puts "Error: Could not find DateLog in BWS file"
    exit 1
  end

  # Parse date (format: "01/28/26 00:00:00" or just "01/28/26")
  date_str = first_row['DateLog'].to_s.split(' ').first
  session_date = Date.strptime(date_str, '%m/%d/%y')
  session_date_str = session_date.strftime("%Y-%m-%d")
  tab_name = session_date.strftime('%m/%d')
  puts "Session date: #{tab_name}"
else
  session_date = Date.strptime(File.basename(INPUT_FILE).split('.').first, '%m-%d-%y')
  session_date_str = session_date.strftime("%Y-%m-%d")
  tab_name = File.basename(INPUT_FILE).split('-')[0,2].join("/")
  puts "Session date: #{tab_name}"
end

# Convert board nums to links
url='https://tcgcloud.bridgefinesse.com/PHPPOSTCGS.php?options=LookupClioBoard&acblno=2909510&date=_DATE_&board=_BOARD_&gamemode=Nite'
csv_data[0] << "Could we do better?"
csv_data[1..-1].each do |row|
  bnum = row[0].to_i
  row[0] = "=HYPERLINK(\"#{url.sub("_DATE_", session_date_str).sub("_BOARD_", "%02d" % bnum)}\",\"#{bnum}\")"
  # Copy % vs Field to the end
  row << row[4]
end

# Step 3: Set up Google Sheets API
puts "\nStep 3: Connecting to Google Sheets API..."
service = Google::Apis::SheetsV4::SheetsService.new
scopes = ['https://www.googleapis.com/auth/spreadsheets']
service.authorization = Google::Auth::ServiceAccountCredentials.make_creds(
  json_key_io: File.open(CREDENTIALS_PATH),
  scope: scopes
)

# Step 4: Get template sheet info
puts "\nStep 4: Reading template sheet #{TEMPLATE_SHEET_ID}..."
spreadsheet = service.get_spreadsheet(TEMPLATE_SHEET_ID)
template_sheet = spreadsheet.sheets.find { |s| s.properties.title == TEMPLATE_TAB_NAME }

unless template_sheet
  puts "Error: Template tab '#{TEMPLATE_TAB_NAME}' not found in sheet"
  puts "Available tabs: #{spreadsheet.sheets.map { |s| s.properties.title }.join(', ')}"
  exit 1
end

template_sheet_id = template_sheet.properties.sheet_id
puts "Found template tab with ID: #{template_sheet_id}"

# Step 5: Check if tab already exists
existing_tab = spreadsheet.sheets.find { |s| s.properties.title == tab_name }
if existing_tab
  puts "\nWarning: Tab '#{tab_name}' already exists. Overwriting..."
  # Delete existing tab
  delete_request = Google::Apis::SheetsV4::Request.new(
    delete_sheet: Google::Apis::SheetsV4::DeleteSheetRequest.new(
      sheet_id: existing_tab.properties.sheet_id
    )
  )
  batch_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(
    requests: [delete_request]
  )
  service.batch_update_spreadsheet(TEMPLATE_SHEET_ID, batch_request)
  puts "Deleted existing tab"
end

# Step 6: Duplicate template tab
puts "\nStep 6: Duplicating template tab..."
duplicate_request = Google::Apis::SheetsV4::Request.new(
  duplicate_sheet: Google::Apis::SheetsV4::DuplicateSheetRequest.new(
    source_sheet_id: template_sheet_id,
    new_sheet_name: tab_name,
    insert_sheet_index: 0  # Insert at beginning
  )
)

batch_request = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new(
  requests: [duplicate_request]
)

response = service.batch_update_spreadsheet(TEMPLATE_SHEET_ID, batch_request)
new_sheet_id = response.replies.first.duplicate_sheet.properties.sheet_id
puts "Created new tab '#{tab_name}' (ID: #{new_sheet_id})"

# Step 7: Clear old data and write new CSV data
puts "\nStep 7: Writing CSV data to new tab..."

# Figure out where the data starts (assuming row 1 is headers)
data_range = "#{tab_name}!A1:N#{TEMPLATE_MAX_ROW}"

# Clear the data range first
clear_request = Google::Apis::SheetsV4::ClearValuesRequest.new
service.clear_values(TEMPLATE_SHEET_ID, data_range, clear_request)
puts "Cleared data range"

# Write new data
value_range = Google::Apis::SheetsV4::ValueRange.new(
  range: data_range,
  values: csv_data
)

service.update_spreadsheet_value(
  TEMPLATE_SHEET_ID,
  data_range,
  value_range,
  value_input_option: 'USER_ENTERED'
)
puts "Wrote #{csv_data.size} rows to sheet"

# Step 8: Done!
sheet_url = "https://docs.google.com/spreadsheets/d/#{TEMPLATE_SHEET_ID}/edit#gid=#{new_sheet_id}"
puts "\n✅ Success!"
puts "Session: #{tab_name}"
puts "Sheet: #{sheet_url}"
