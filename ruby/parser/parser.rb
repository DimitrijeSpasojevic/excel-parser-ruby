require 'roo'
require 'spreadsheet'

class Column
  attr_accessor :name, :fields, :table

  def initialize(name, fields, table)
      @name = name
      @fields = fields
      @table = table
  end

  def sum
      fields_sum = 0
      @fields.each do |field|
          fields_sum += field
      end

      fields_sum
  end

  def [](nthEl)
    fields[nthEl]
  end

  def sum
    fields.sum
  end
  
  def make_cell_methods
    @fields.each do |cell|
        self.define_singleton_method("find_#{cell}") do
            index = -1
            @table.rows.each_with_index do |row, i|
                if i == 0
                    row.each_with_index do |row_cell, j|
                        if row_cell == @name
                            index = j
                        end
                    end
                else
                    row.each_with_index do |row_cell, j|
                        if row_cell == cell and j == index
                            return row
                        end
                    end
                end
            end
        end
    end
  end
end


class Table
    attr_accessor :name, :columns, :rows
    include Enumerable

    def initialize(name)
        @name = name
        @columns = Array.new
        @rows = Array.new
    end

    def getTable()
      return rows
    end

    def row(nth)
      rows[nth]
    end 

    def [](columName)
      @columns.each do |col|
        if col.name == columName
          return col
        else
          return "Ne postoji"
        end
      end
    end  

    def init_columns
      table_columns = @rows.transpose
      
      table_columns.each do |col|
          current_column = Column.new(col[0], col[1..-1], self)
          @columns << current_column
      end

      @columns.each do |col|
          col_name = col.name

          self.define_singleton_method("#{col_name}") do
              @columns.each do |column|
                  if column.name == col_name
                      return col
                  end
              end
          end
          col.make_cell_methods
      end
    end

    def each(&block)
      @rows.each do |row|
        @columns.each do |col|
          block.call(rows[rows.find_index(row)][columns.find_index(col)])
        end
      end
    end
end

def initTables(path)
  tables = Array.new()
  if path.end_with?(".xlsx")     
    workbook = Roo::Spreadsheet.open(path, {:expand_merged_ranges => true})
    worksheets = workbook.sheets
    #puts "Found #{worksheets.count} worksheets"
    worksheets.each do |worksheet|
      #puts "Reading: #{worksheet}"
      num_rows = 0
      tables.append(Table.new("Table_#{worksheet}"))
      workbook.sheet(worksheet).each_row_streaming do |row|
        row_cells = row.map { |cell| cell.value }
        row_cells_f = row.map { |cell| cell }
        #print row_cells_f
        tables[-1].rows.append(row_cells) unless row_cells_f.to_s.include? "@formula=\"SUBTOTAL" or row_cells_f.to_s.include? "@formula=\"TOTAL"
        num_rows += 1
      end
      tables[-1].init_columns
      #puts "Read #{num_rows} rows" 
    end
  else
    # Note: spreadsheet only supports .xls files (not .xlsx)
    workbook = Spreadsheet.open(path, {:expand_merged_ranges => true})
    worksheets = workbook.worksheets
    #puts "Found #{worksheets.count} worksheets"

    worksheets.each do |worksheet|
      #puts "Reading: #{worksheet.name}"
      num_rows = 0
      tables.append(Table.new("Table_#{worksheet}"))
      worksheet.rows.each do |row|
        row_cells = row.to_a.map{ |v| v.methods.include?(:value) ? v : v }
        unless row_cells.length() == 0 or row_cells.to_s.include? "Formula"
          row_cells.compact!
          tables[-1].rows.append(row_cells)
          #print row_cells, "\n"
        end
        num_rows += 1
      end
      tables[-1].init_columns
      #puts "-----------------------Gotova jedan tabela---------------------------"
    end
  end
  tables
end


def plus(t1,t2)
  if t1.rows[0] != t2.rows[0]
    raise "Headers are not the same!"
  else
    table = Table.new("Zbir")
    table.rows += t1.rows
    table.rows += t2.rows[1..-1]
    table.init_columns
  end
  table
end

def minus(t1,t2)
  table = t1.dup
  if t1.rows[0] != t2.rows[0]
    raise "Headers are not the same!"
  else
    t2.rows[1..-1].each do |row_from_2|
      table.rows[1..-1].each do |row_from_1|
        if row_from_2 == row_from_1
          table.rows.delete(row_from_1)
        end  
      end
    end  
  end
  table
end