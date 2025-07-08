#!/usr/bin/env ruby

require_relative 'borrowed_methods.rb'
require_relative 'new_methods.rb'
require_relative 'borrowed_classes.rb'

def add_year(sheet_id, start)
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
  new_start_letter = biject(last_col + 1)
  border_range = "#{sheet.title}!#{new_start_letter}1:#{new_start_letter}#{last_row}"

  # UNGROUP LAST YEAR
  # Have to do this before adding columns,
  # or it will group the last two years together.
  multiple_batch_update(delete_column_groups(sheet_id, last_col, last_col - 12))

  list_of_reqs = []
  # ADD COLUMNS
  list_of_reqs += add_columns(sheet_id, last_col, 12)
  # UNMERGE TOP CELLS
  list_of_reqs += unmerge_cells(sheet_id, 0, 2, start, last_col)
  # MERGE TOP CELLS + NEW
  list_of_reqs += merge_cells(sheet_id, 0, 2, start, last_col + 12)
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
    2.times do |i|
      increment = i * 12
      list_of_reqs.push(
        add_dimension_group: {
        range: {
          dimension: "COLUMNS",
          sheet_id: sheet_id,
          start_index: last_col - 11 + increment,
          end_index: last_col + increment
        },
        }
      )
  end

  multiple_batch_update(list_of_reqs)
  save("#{sheet.title}!#{new_start_letter}3", months_array)
end

# YEAR WITH QUARTERS
def add_year_with_quarters(sheet_id, start)
  next_year = Time.now.year + 1
  months = [
    "Jan", "Feb", "Mar", "Q1",
    "Apr", "May", "Jun", "Q2",
    "Jul", "Aug", "Sep", "Q3",
    "Oct", "Nov", "Dec", "Q4"
  ]
  months_with_year = months.map { |month| "#{month} #{next_year}" }
  months_array = [months_with_year]
  sheet = $list.find { |s| s.sheet_id == sheet_id }
  last_col = sheet.grid_properties.column_count
  last_row = sheet.grid_properties.row_count
  new_start_letter = biject(last_col + 1)
  border_range = "#{sheet.title}!#{new_start_letter}1:#{new_start_letter}#{last_row}"

  # UNGROUP LAST YEAR
  # Have to do this before adding columns,
  # or it will group the last two years together.
  multiple_batch_update(delete_column_groups(sheet_id, last_col, last_col - 16))

  list_of_reqs = []
  # ADD COLUMNS
  list_of_reqs += add_columns(sheet_id, last_col, 16)
  # CLEAR BACKGROUND
  list_of_reqs += remove_background_color(sheet_id, 3, last_row, last_col, last_col + 3)
  list_of_reqs += remove_background_color(sheet_id, 3, last_row, last_col + 4, last_col + 7)
  list_of_reqs += remove_background_color(sheet_id, 3, last_row, last_col + 8, last_col + 11)
  list_of_reqs += remove_background_color(sheet_id, 3, last_row, last_col + 12, last_col + 15)
  # UNMERGE TOP CELLS
  list_of_reqs += unmerge_cells(sheet_id, 0, 2, start, last_col)
  # MERGE TOP CELLS + NEW
  list_of_reqs += merge_cells(sheet_id, 0, 2, start, last_col + 16)
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
    2.times do |i|
      increment = i * 16
      list_of_reqs.push(
        add_dimension_group: {
        range: {
          dimension: "COLUMNS",
          sheet_id: sheet_id,
          start_index: last_col - 15 + increment,
          end_index: last_col + increment
        },
        }
      )
  end
  
  multiple_batch_update(list_of_reqs)
  save("#{sheet.title}!#{new_start_letter}3", months_array)
end

# RUN DEMO
$link = '1wb6QmfpyXGEv9cOAVNEdUMvggGVZjyTfMFTqlWSAz6Q'
$list = google.get_spreadsheet($link).sheets.map {|s| s.properties }
black_list = [1972370723, 1538279974]
start = 1
$list.each do |sheet|
  if sheet.sheet_id == 1538279974
    add_year_with_quarters(sheet.sheet_id, start)
    copy_formulas(sheet.sheet_id, 16)
  elsif !black_list.include?(sheet.sheet_id)
    add_year(sheet.sheet_id, start)
    copy_formulas(sheet.sheet_id, 12)
  end
end