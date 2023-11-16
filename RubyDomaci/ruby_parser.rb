require 'roo'
require 'spreadsheet'


class Column
    attr_accessor :header, :cells, :table #generise getter i setter


    def initialize(header, cells, table)
      @header = header
      @cells = cells
      @table = table
    end


    #dodavanje dinamickih metoda da se za svaku celiju iz kolone vrati red kom pripada
    def dynamic_cell_methods
      @cells.each do |cell|
        self.define_singleton_method("cell_#{cell}") do
          column_index = -1
          #prvo nadjemo ogovarajuci index kolone koji se poklapa sa hederom(@header)
          @table.rows.each_with_index do |row, i|
            if i == 0
              row.each_with_index do |row_cell, j|
                if row_cell == @header
                  column_index = j
                end
              end
            else
              row.each_with_index do |row_cell, j|
                if row_cell == cell and j == column_index
                  return row
                end
              end
            end
          end
        end
      end
    end


    #racunanje sume za kolonu
    def column_sum
      cells_sum = 0
      @cells.each do |cell|
        cells_sum += cell
      end

      return cells_sum
    end

    def avg
      sum = @cells.reduce(0) { |acc, cell| acc + cell.to_f }
      count = @cells.count
      sum / count.to_f
    end

    def map(&block)
      @cells.map(&block)
    end

    def select(&block)
      @cells.select(&block)
    end

    def reduce(initial, &block)
      @cells.reduce(initial, &block)
    end
end


class Table
  attr_accessor :rows, :columns, :excel_file
  include Enumerable


  def initialize(path, sheet)
    @rows = []
    @columns = []

    self.read_table(path, sheet)
  end

  def read_table(path, sheet)
    if path.end_with?(".xlsx")
      @excel_file = Roo::Spreadsheet.open(path, {:expand_merged_ranges => true})
      index = 0;
      index_of_rows_for_delete = Array.new

      @excel_file.sheet(sheet).each_row_streaming do |row|
        formulas = row.map {|cell| cell}

        if formulas.to_s.include? "@formula=\"SUBTOTAL" or formulas.to_s.include? "@formula=\"TOTAL"
          index_of_rows_for_delete << index
        end

        index += 1
      end


      @excel_file.sheet(sheet).each_with_index do |row, i|
        if index_of_rows_for_delete.include? i
          next
        end

        if row.all? { |x| x.nil? }
          next
        end

        row.each_with_index do |cell_data, i|
          if cell_data == nil
            row[i] = 0
          end
        end

        @rows << row
      end

      self.add_columns

    elsif path.end_with?(".xls")
      @excel_file = Spreadsheet.open(path)

      for row in @excel_file.worksheet(sheet)
        if row.all? { |x| x.nil? }
          next
        end

        insert = true
        row.each_with_index do |cell_data, i|
          if cell_data.to_s.include? "Formula"
            insert = false
          end

          if cell_data == nil
            row[i] = 0
          end
        end

        if insert
          @rows << row
        end

      end

      #Brisemo sve prazne kolone sa leve strane tabele
      transposed_table = @rows.transpose
      transposed_table.delete_if do |col|
        if col.all? { |x| x == 0 }
          true
        end
      end

      #Odsecamo decimale brojevima koji to imaju
      # @rows = transposed_table.transpose
      # @rows.each_with_index do |row, i|
      #   #Ako je decimala, odsecemo je
      #   row.each_with_index do |cell, j|
      #     if cell !~ /\D/
      #       @rows[i][j] = cell.round
      #     end
      #   end
      # end
      @rows = transposed_table.transpose
      @rows.each_with_index do |row, i|
        # Ako je decimala, zaokružujemo je
        row.each_with_index do |cell, j|
          if cell.is_a?(Float)
            @rows[i][j] = cell.to_i
          end
        end
      end

      self.add_columns
    else
      puts "Wrong extention!"
    end
  end

  def add_columns
    table_columns = @rows.transpose

    table_columns.each do |col|
      @columns << Column.new(col[0], col[1..-1], self)
    end

    @columns.each do |col|
      header = col.header

      self.define_singleton_method("#{header}") do
        @columns.each do |column|
          if column.header == header
            return col
          end
        end
      end

      col.dynamic_cell_methods
    end
  end


  def [](header)
    @columns.each do |col|
      if col.header == header
        return col.cells
      end
    end
  end


  def []=(header, index, value)
    column = @columns.find { |col| col.header == header }
    column.cells[index] = value if column
  end


  def each(&block)
    @rows.each do |row|
      block.call(row)
    end
  end


  def row(index)
    return @rows[index]
  end


  def print_table
    for row in @rows
      print row, "\n"
    end
  end


  def +(union_table)
    if @rows[0] != union_table.rows[0]
      return
    end

    union_table_rows = union_table.rows
    #Dodamo nove redove iz druge tabele
    union_table_rows.each_with_index do |row, i|
      #Preskacemo hedere
      if i == 0
        next
      end
      @rows << row
    end

    #Obrisemo sve kolone i ponovo ih inicijalizujemo, kako bi svaka celija dobila svoju metodu
    @columns = @columns.clear
    self.add_columns
  end


  def -(difference_table)
    if @rows[0] != difference_table.rows[0]
      return
    end

    difference_table_rows = difference_table.rows
    difference_table_rows.each_with_index do |row, i|
      if i == 0
        next
      end

      @rows.delete_if do |table_row|
        if table_row == row
          true
        end
      end
    end

    @columns = @columns.clear
    self.add_columns
  end

end





