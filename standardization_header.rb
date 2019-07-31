require 'rubyXL'
require 'rubyXL/convenience_methods/cell'
plans = ['2005', '2006', '2007', '2008',
         '2009', '2010', '2011', '2012',
         '2013', '2014', '2015', '2016',
         '2017', '2018', '2019']

plans.each_with_index do |age, index|

  nome_planilha = "_CONTROLE_DE_CORRESPONDENCIA_#{age}.XLSX"

  workbook = RubyXL::Parser.parse("planilhas/DARES_PADRONIZADAS/#{nome_planilha}")

  worksheets = workbook.worksheets
  worksheets.each do |worksheet|
    num_rows = 0
    if worksheet.sheet_name.to_s == "DARES"
      worksheet.each_with_index do |row, i|
        if i == 0
          row.cells.each_with_index.map{ |cell, j|
            unless cell.nil? && j < 30
              cell.change_contents(cell.value.to_s.downcase.gsub("á", "a")
                                                           .gsub("à", "a")
                                                           .gsub("ã", "a")
                                                           .gsub("â", "a")
                                                           .gsub("è", "e")
                                                           .gsub("é", "e")
                                                           .gsub("ẽ", "e")
                                                           .gsub("ê", "e")
                                                           .gsub("í", "i")
                                                           .gsub("ì", "i")
                                                           .gsub("ó", "o")
                                                           .gsub("ò", "o")
                                                           .gsub("ô", "o")
                                                           .gsub("õ", "o")
                                                           .gsub("ú", "u")
                                                           .gsub("ù", "u")
                                                           .gsub("ũ", "u")
                                                           .gsub("û", "u")
                                                           .gsub("ç", "c")
                                                           .gsub("-", "")
                                                           .gsub(" ", "_")
                                                           .gsub("º", "").upcase)
            end
          }
        end
      end
    end
  end

  workbook.save("./output/#{nome_planilha}")
  puts "#{nome_planilha} DONE!"
end
