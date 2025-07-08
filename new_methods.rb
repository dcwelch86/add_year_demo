require_relative 'borrowed_methods.rb'

def add_columns(sheet_id, start_index, number_of_columns)
  reqs = []
  reqs.push(
    insert_dimension: {
      range: {
        sheet_id: sheet_id,
        dimension: "COLUMNS",
        start_index: start_index,
        end_index: start_index + number_of_columns
      },
      inherit_from_before: true
    }
  )
  return reqs
end

def merge_cells(sheet_id, start_row, end_row, start_column, end_column)
  reqs = []
  reqs.push(
    merge_cells: {
      range: {
        sheet_id: sheet_id,
        start_row_index: start_row,
        end_row_index: end_row,
        start_column_index: start_column,
        end_column_index: end_column
      },
      merge_type: "MERGE_ALL"  # Options: MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS
    }
  )
  return reqs
end

def unmerge_cells(sheet_id, start_row, end_row, start_column, end_column)
  reqs = []
  reqs.push(
    unmerge_cells: {
      range: {
        sheet_id: sheet_id,
        start_row_index: start_row,
        end_row_index: end_row,
        start_column_index: start_column,
        end_column_index: end_column
      }
    }
  )
  return reqs
end

def delete_column_groups(sheet_id, end_index, start_index = 1)
  reqs = []
  reqs.push(
      delete_dimension_group: {
          range: {
            sheet_id: sheet_id,
            dimension: "COLUMNS",
            start_index: start_index,
            end_index: end_index
          },
      }
  )
  return reqs
end

def remove_background_color(sheet_id, start_row_index, end_row_index, start_column_index, end_column_index)
  reqs = []
  reqs.push(
    repeat_cell: {
      range: {
        sheet_id: sheet_id,
        start_row_index: start_row_index,  
        end_row_index: end_row_index,
        start_column_index: start_column_index,
        end_column_index: end_column_index
      },
      cell: {
        user_entered_format: {
          background_color: {
            red: 1.0,
            green: 1.0,
            blue: 1.0
          }  # This clears the background color
        }
      },
      fields: 'userEnteredFormat.backgroundColor'
    }
  )
  return reqs
end

def copy_paste(source, destination)
  reqs = []
  reqs.push(
      {
        copy_paste: {
          source: {
              sheet_id: range(source)[:sheet_id],
              start_row_index: range(source)[:start_row_index],
              end_row_index: range(source)[:end_row_index],
              start_column_index: range(source)[:start_column_index],
              end_column_index: range(source)[:end_column_index]
          },
          destination: {
              sheet_id: range(destination)[:sheet_id],
              start_row_index: range(destination)[:start_row_index],
              end_row_index: range(destination)[:end_row_index],
              start_column_index: range(destination)[:start_column_index],
              end_column_index: range(destination)[:end_column_index]
          },
          paste_type: "PASTE_FORMULA",
          paste_orientation: "NORMAL"
        }
      }
  )
  begin
    retries ||= 0
    resp = google.batch_update_spreadsheet($link, { requests: reqs })
    resp
  rescue => e
    retries += 1
    sheet_error(e)
    retry if retries < 3
  end
end

def delete_non_formulas(area)
  vals = google.get_spreadsheet_values($link, area, value_render_option: 'FORMULA').values
  puts "No values in #{area}." if vals.nil?
  return if vals.nil? 

  save_cell = area.split(":")[0]
  rows = vals.map do |sub_array|
    sub_array.map { |value| value.is_a?(String) ? value : "" }
  end
  save(save_cell, rows)
end

def copy_formulas(sheet_id, number_of_columns)
  sheet = $list.find { |s| s.sheet_id == sheet_id }
  # Source for copying formulas
  source = "#{sheet.title}!#{biject(sheet.grid_properties.column_count)}4:#{biject(sheet.grid_properties.column_count)}#{sheet.grid_properties.row_count}"
  destination = "#{sheet.title}!#{biject(sheet.grid_properties.column_count + 1)}4:#{biject(sheet.grid_properties.column_count + number_of_columns)}#{sheet.grid_properties.row_count}"

  copy_paste(source, destination)
  delete_non_formulas(destination)
end