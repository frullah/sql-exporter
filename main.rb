# frozen_string_literal: true

require 'oci8'
require 'fast_excel'

group_by = 'Tahun Pajak'
sql_initialize = File.read('./sql-initialize.sql')
sql = File.read('./sql.sql')
book = FastExcel.open("./output/result-#{DateTime.now}.xlsx")
sheets = {}
connection = OCI8.new('iprotax', 'iprotax', '172.17.5.20/iprotax')
connection.exec(sql_initialize)
cursor = connection.parse(sql)
cursor.exec

def add_row(book:, sheets:, row:, group: nil)
  group_value = row[group]
  sheets[group_value] ||= begin
    sheet = book.add_worksheet(group_value)
    sheet.auto_width = true
    sheet.append_row(row.keys)
    sheet
  end
  sheet = sheets[group_value]
  sheet.append_row(row.values)
end

puts "Row count : #{cursor.row_count}"
cursor.fetch_hash do |row|
  add_row(book: book, sheets: sheets, row: row, group: group_by)
end

book.close
#   puts row
# end
