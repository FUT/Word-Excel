require 'spreadsheet'

def pp(*options)
  puts '=' * 180, options, "\n\n"
end


Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet.open 'xls.xls'

pp "Worksheets:\n#{book.worksheets.map(&:name).join("\n")}\n\n"

sheet1 = book.worksheet 0
sheet2 = book.worksheet 'second'

pp sheet1.map{|row| row.join "\t\t"}.join "\n"
pp sheet2.map{|row| row.join "\t\t"}.join "\n"

pp sheet1.map{|row| row.formats.map{|f| "font=#{f.font.name};" if f}.join "\t\t"}.join "\n"
pp sheet1.map{|row| row.formats.map{|f| "size=#{f.font.size};" if f}.join "\t\t"}.join "\n"

#CREATE NEW ONE
new_book = Spreadsheet::Workbook.new
new_sheet = new_book.create_worksheet :name => 'New Worksheet '

#INSERT DATA
new_sheet.row(0).concat %w{Name Country Acknowlegement}
new_sheet[1,0] = 'Japan'
row = new_sheet.row(1)
row.push 'Creator of Ruby'
row.unshift 'Yukihiro Matsumoto'
new_sheet.row(2).replace [ 'Daniel J. Berger', 'U.S.A.',
                        'Author of original code for Spreadsheet::Excel' ]
new_sheet.row(3).push 'Charles Lowe', 'Author of the ruby-ole Library'
new_sheet.row(3).insert 1, 'Unknown'
new_sheet.update_row 4, 'Hannes Wyss', 'Switzerland', 'Author'

#FORMATTING
new_sheet.row(0).height = 18

format = Spreadsheet::Format.new :color => :blue,
                                 :weight => :bold,
                                 :size => 18
new_sheet.row(0).default_format = format

bold = Spreadsheet::Format.new :weight => :bold
4.times{|x| new_sheet.row(x + 1).set_format(0, bold)}

#SAVE
new_book.write 'xls_new.xls'

