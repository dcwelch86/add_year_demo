#!/usr/bin/env ruby

require_relative 'borrowed_methods.rb'
require_relative 'new_methods.rb'
require_relative 'borrowed_classes.rb'

# ONLY WORKS WHEN DOING A SINGLE SHEET

def add_year(sheet_id)
  next_year = Time.now.year + 1
  months = [
    "Jan", "Feb", "Mar", "Apr",
    "May", "Jun", "Jul", "Aug",
    "Sep", "Oct", "Nov", "Dec"
  ]
  months_with_year = months.map { |month| "#{month} #{next_year}" }
  months_array = [months_with_year]
  sheet = $list.find { |s| s.sheet_id == sheet_id }
  last_col = sheet.grid_properties.column_count
  last_row = sheet.grid_properties.row_count
  cur_end_letter = biject(last_col)
  new_start_letter = biject(last_col + 1)
  new_end_letter = biject(last_col + 12)
  border_range = "#{sheet.title}!#{new_start_letter}1:#{new_start_letter}#{last_row}"
  num_years = ((last_col/12) + 1).floor
  group_start = (last_col % 12) + 1
  group_end = group_start + 11
  # Source for copying formulas
  source = "#{sheet.title}!#{cur_end_letter}4:#{cur_end_letter}#{last_row}"
  destination = "#{sheet.title}!#{new_start_letter}4:#{new_end_letter}#{last_row}"

  # UNGROUP ALL
  # Have to do this before adding columns,
  # or it will group the last two years together.
  multiple_batch_update(delete_column_groups(sheet_id, last_col))

  list_of_reqs = []
  # ADD COLUMNS
  list_of_reqs += add_columns(sheet_id, last_col, 12)
  # UNMERGE TOP CELLS
  list_of_reqs += unmerge_cells(sheet_id, 0, 2, group_start - 1, last_col)
  # MERGE TOP CELLS + NEW
  list_of_reqs += merge_cells(sheet_id, 0, 2, group_start - 1, last_col + 12)
  # ADD BORDER
  list_of_reqs +=
    add_border(
      area: border_range,
      top: nil,
      bottom: nil,
      left: CellBorder.new(style: "SOLID", color: "#000000"),
      right: nil,
      inner_horizontal: nil,
      inner_vertical: nil 
    )
  # CREATE YEAR GROUP
  num_years.times do |i|
      increment = i * 12
      list_of_reqs.push(
        add_dimension_group: {
        range: {
          dimension: "COLUMNS",
          sheet_id: sheet_id,
          start_index: group_start + increment,
          end_index: group_end + increment
        },
        }
      )
  end
  # COLLAPSE ALL BUT LAST TWO YEARS
  (num_years - 2).times do |i|
      increment = i * 12
      list_of_reqs.push(
        update_dimension_group: {
          dimension_group: {
            range: {
              sheet_id: sheet_id,
              dimension: "COLUMNS",
              start_index: group_start + increment,
              end_index: group_end + increment
            },
            depth: 1,
            collapsed: true
          },
          fields: "collapsed"
        }
      )
  end
  multiple_batch_update(list_of_reqs)
  save("#{sheet.title}!#{new_start_letter}3", months_array)
  copy_paste(source, destination)
  delete_non_formulas(destination)
end