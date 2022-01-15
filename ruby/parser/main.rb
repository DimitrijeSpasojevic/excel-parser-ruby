require_relative 'parser'
#1
tables = initTables('../tables/table1.xlsx');

t1 = Array.new()
t1 = tables[0].getTable
#2
#p t1[1][2]

t = tables[0]
t2 = tables[1]

#3
#p t.row(0)[4]

#4
#t.each { |el| puts "* #{el}" }

#5 merge table2 primer

#6
#puts t["Kolona1"].fields
#puts "Drugi elem 1. kolone"
#puts t["Kolona1"][1]

#7
#puts t.Kolona1.fields

#7.1 sum in column
#puts t.Kolona1.sum
#7.2
#puts t.Kolona2.find_22


#8
# t.rows.each do |row|
#    p row
#    p "-----------"
# end

#9
#tt = plus t,t2
#tt.rows.each do |row|
#    p row
#    p "-----------"
#end

#10
# tt = minus t,t2
# tt.rows.each do |row|
#    p row
#    p "-----------"
# end

# t.rows.each do |row|
#    p row
#    p "-----------"
# end