require 'docx'

doc = Docx::Document.open('tables.docx')

first_table = doc.tables[0]
puts first_table.row_count
puts first_table.column_count
puts first_table.rows[0].cells[0].text
puts first_table.columns[0].cells[0].text

doc.tables.each do |table|
  table.rows.each do |row|
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table.columns.each do |column|
    column.cells.each do |cell|
      puts cell.text
    end
  end
end