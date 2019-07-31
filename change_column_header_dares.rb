require 'rubyXL'
require 'rubyXL/convenience_methods/cell'
plans = ['2005', '2006', '2007', '2008',
         '2009', '2010', '2011', '2012',
         '2013', '2014', '2015', '2016',
         '2017', '2018', '2019']

plans.each_with_index do |age, index|

  nome_planilha = "_CONTROLE_DE_CORRESPONDENCIA_#{age}.XLSX"

  # workbook = RubyXL::Parser.parse("input/DARES_PADRONIZADAS/#{nome_planilha}")
  workbook = RubyXL::Parser.parse("input/dares/#{nome_planilha}")

  worksheets = workbook.worksheets
  worksheets.each do |worksheet|
    num_rows = 0
    if worksheet.sheet_name.to_s == "DARES"
      worksheet.each_with_index do |row, i|
        if i == 0
          row.cells.each_with_index.map{ |cell, j|
            if !cell.nil? && j < 7 && cell.value.to_s == 'DATA'
              cell.change_contents('DATA_DO_DOCUMENTO')
            elsif !cell.nil? && j > 7 && (cell.value.to_s == 'DATA' || cell.value.to_s == 'DATA_DE_PROTOCOLO')
              cell.change_contents('DATA_PROTOCOLO')
            elsif !cell.nil? && cell.value.to_s == 'DATA_DE_PROTOCOLO_SEI'
              cell.change_contents('DATA_PROTOCOLO_SEI')
            elsif !cell.nil? && cell.value.to_s == 'TIPO'
              cell.change_contents('TIPO_DO_DOCUMENTO')
            elsif !cell.nil? && cell.value.to_s == 'N'
              cell.change_contents('NUMERO_DE_DOCUMENTO')
            end
          }
        end
      end
    end
  end

  workbook.save("./output/DARES_PADRONIZADAS/#{nome_planilha}")
  puts "#{nome_planilha} DONE!"
end
