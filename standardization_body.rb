require 'rubyXL'
require 'rubyXL/convenience_methods/cell'

MAX_LINES = 57
MAX_COLUMNS = 21
CANCELED_INDEX = 21

nome_planilha = "_Controle de Ofícios 2019.xlsx"

workbook = RubyXL::Parser.parse("./planilhas/#{nome_planilha}")

worksheets = workbook.worksheets
puts worksheets[0].sheet_name
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  num_rows = 0
  if worksheet.sheet_name.to_s == "OFICIOS"
    worksheet.each_with_index do |row, i|
      if i <= MAX_LINES
        num_rows += 1
        columns = 1
        row.cells.each_with_index.map{ |cell, j|
          if !cell.nil? && j <= MAX_COLUMNS
            if (cell.value.to_s.gsub(" ", "") == "" || cell.value.nil? ||
                (cell.value.to_s.upcase == "SEM PRAZO") ||
                (cell.value.to_s.upcase == "S/C") ||
                (cell.value.to_s.upcase == "S/N") ||
                (cell.value.to_s.upcase == "NA") ||
                (cell.value.to_s.upcase == "N/A") ||
                (cell.value.to_s.upcase == "EM BRANCO") ||
                (cell.value.to_s.gsub("ã", "a").upcase == "NAO SE APLICA") ||
                (cell.value.to_s.gsub("ú", "u").upcase == "SEM NUMERO") ||
                (cell.value.is_a?(Integer) && cell.value.to_i <= 0) ||
                (cell.value.is_a?(Float) && cell.value.to_i <= 0))

              cell.change_contents("N/A")
            elsif i > 0 &&
                  (cell.value.to_s.upcase.gsub(" ", "") == "CANCELADA" ||
                  cell.value.to_s.upcase.gsub(" ", "") == "CANCELADO")

              worksheet[i][CANCELED_INDEX].change_contents("SIM")
            end

            # Padroniza datas
            # if (j == 4 || j == 9) && i > 0 &&
            #   cell.value.to_s.gsub(" ", "").upcase != "N/A" &&
            #   cell.value.to_s.gsub(" ", "").upcase != "SEM PRAZO" &&
            #   cell.value.to_s.gsub(" ", "").upcase != "CANCELADA" &&
            #   cell.value.to_s.gsub(" ", "").upcase != "CANCELADO" &&
            #   cell.value.to_s.gsub("á", "a").gsub("ã","a").upcase != "NAO SERA USADA"
            #
            #   cell.change_contents(cell.value.to_s + '/2006')
            # end
          end
        }
      end
    end
  end
  puts "Read #{num_rows} rows"
end

workbook.save("./output/#{nome_planilha.gsub(" ", "_").gsub("í","i").upcase}")
