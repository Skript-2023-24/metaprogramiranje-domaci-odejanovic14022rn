require_relative 'ruby_parser'

puts "TABLE 1: XLSX"
t1 = Table.new('./t1.xlsx', 'Sheet1')
t1.print_table
print "\n"

print "Headers: ", t1.row(0), "\n"
print "First row: ", t1.row(1), "\n"
print "First row, third cell: ", t1.row(1)[2], "\n"
print "Third column: ", t1["Third"], "\n"
print "Third column, second cell: ", t1["Third"][2], "\n"

t1_tc = t1.Third
print "Third column - cells: ", t1_tc.cells, "\n"
print "Third column - cells sum: ", t1_tc.column_sum, "\n"
print "Third column - cell's row: ", t1_tc.cell_10, "\n\n"

print "EACH\n"
t1.each {|row| puts "ROW - #{row}"}


puts "\n", "TABLE 1 + TABLE 2:"
t2 = Table.new('./t2.xlsx', 'Sheet1')
t1 + t2
t1.print_table

puts "\n", "TABLE 1 - TABLE 2:"
t1 - t2
t1.print_table


###########################################################################


puts "\n\n", "TABLE 3: XLS"
t3 = Table.new('./t3.xls', 'Sheet1')
t3.print_table
print "\n"

print "Headers: ", t3.row(0), "\n"
print "First row: ", t3.row(1), "\n"
print "First row, third cell: ", t3.row(1)[2], "\n"
print "Third column: ", t3["Third"], "\n"
print "Third column, second cell: ", t3["Third"][2], "\n"

t3_tc = t3.Third
print "Third column - cells: ", t3_tc.cells, "\n"
print "Third column - cells sum: ", t3_tc.column_sum, "\n"
print "Third column - cell's row: ", t3_tc.cell_10, "\n\n"


puts "\n", "TABLE 3 + TABLE 4:"
t4 = Table.new('./t4.xls', 'Sheet1')
t3 + t4
t3.print_table

puts "\n", "TABLE 3 - TABLE 2:"
t3 - t4
t3.print_table