require 'rubyXL'
require 'rubyXL/convenience_methods/cell'
plans = ['1998', '1999', '2000', '2001',
         '2002', '2003', '2004', '2005',
         '2007', '2008', '2009', '2010',
         '2011', '2012', '2013', '2014',
         '2015', '2016', '2017', '2018',
         '2019']

plans.each_with_index do |age, index|

  nome_planilha = "_CONTROLE_DE_OFICIOS_#{age}.XLSX"

  # workbook = RubyXL::Parser.parse("planilhas/DARES_PADRONIZADAS/#{nome_planilha}")
  workbook = RubyXL::Parser.parse("planilhas/oficios/#{nome_planilha}")

  worksheets = workbook.worksheets
  worksheets.each do |worksheet|
    num_rows = 0
    if worksheet.sheet_name.to_s == "OFICIOS"
      worksheet.each_with_index do |row, i|
        if i == 0
          row.cells.each_with_index.map{ |cell, j|
            if !cell.nil? && j < 30 && cell.value.to_s == 'DATA_DO_OFICIO'
              cell.change_contents('DATA_DO_DOCUMENTO')
              puts "DATA_DO_DOCUMENTO = #{age}"
            elsif !cell.nil? && j < 30 && cell.value.to_s == 'N_DOCUMENTO'
              cell.change_contents('NUMERO_DE_DOCUMENTO')
              puts "NUMERO_DE_DOCUMENTO = #{age}"
            elsif !cell.nil? && j < 30 && cell.value.to_s == 'DATA_DE_RECEBIMENTO'
              cell.change_contents('DATA_RECEBIMENTO')
              puts "DATA_RECEBIMENTO = #{age}"
            end
          }
        end
      end
    end
  end

  workbook.save("./output/OFICIOS_PADRONIZADOS/#{nome_planilha}")
  puts "#{nome_planilha} DONE!"
end
