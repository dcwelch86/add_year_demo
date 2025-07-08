#!/usr/bin/env ruby

require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'
require 'google/apis/drive_v3'

OOB_URI          = 'urn:ietf:wg:oauth:2.0:oob'.freeze
APPLICATION_NAME = 'Project_Test'.freeze
CREDENTIALS_PATH = File.path('oauth_2_client_id.json').freeze
TOKEN_PATH       = File.path('token.yaml').freeze
SCOPE            = Google::Apis::SheetsV4::AUTH_SPREADSHEETS

def google
  $google ||= begin
    connect = -> {
      idno = Google::Auth::ClientId.from_file(CREDENTIALS_PATH)
      repo = Google::Auth::Stores::FileTokenStore.new(file: TOKEN_PATH)
      auth = Google::Auth::UserAuthorizer.new(idno, SCOPE, repo)
      user = 'default'
      credentials = auth.get_credentials(user) || begin
        href = auth.get_authorization_url(base_url: OOB_URI)
        puts "Open the following URL and paste the code here:\n" + href
        code = gets
        credentials = auth.get_and_store_credentials_from_code(
          user_id: user, code: code, base_url: OOB_URI
        )
        credentials
      end
    }
    $google = Google::Apis::SheetsV4::SheetsService.new
    $google.client_options.application_name = APPLICATION_NAME
    $google.authorization = connect.call
    $google
  end
end

def sheet_error(e)
  puts "Encountered Error, Sleeping for 30 seconds and trying again"
  p e 
  sleep(2)
end

def sheet_id(obj)
  case obj
    when /^#(\d+)$/ then $list[$1.to_i - 1].sheet_id
    when Integer    then obj
    else $list.first_result {|item| item.sheet_id if item.title == obj}
  end
end

def sheet_name(obj)
  case obj
    when /^#(\d+)$/ then $list[$1.to_i - 1].title
    when Integer    then $list.first_result {|item| item.title if item.sheet_id == obj}
    else obj
  end
end

def save(area, rows, log=true)
  begin
    retries ||= 0
    area.sub!(/^(#\d+)(?=!)/) {|num| sheet_name(num)}
    gasv = Google::Apis::SheetsV4::ValueRange.new(range: area, values: rows)
    done = google.update_spreadsheet_value($link, area, gasv, value_input_option: "USER_ENTERED")
    done
  rescue => e
    retries += 1
    sheet_error(e)
    retry if retries < 3
  end
  puts "#{rows.count} cells updated." if log
  rows.count
end

def save!(area, rows, log=false)
  save(area, rows, log)
end

def biject(x) # a:1, z:26, aa:27, az:52, ba:53, aaa:703
  case x
  when String
    x.each_char.inject(0) {|n,c| (n * 26) + (c.ord & 31) }
  when Integer
    s = []
    s << (((x -= 1) % 26) + 65).chr && x /= 26 while x > 0
    s.reverse.join
  end
end

def range(area)
  sh, rc = area.split('!', 2); rc, sh = sh, nil if sh.nil?
  as, ae = rc.split(':', 2); ae ||= as
  cs, rs = as.split(/(?=\d)/, 2); cs = biject(cs) - 1; rs = rs.to_i - 1
  ce, re = ae.split(/(?=\d)/, 2); ce = biject(ce) - 1; re = re.to_i - 1
  {
    sheet_id:           sh ? sheet_id(sh) : nil,
    start_column_index: cs,
    start_row_index:    rs,
    end_column_index:   ce + 1,
    end_row_index:      re + 1,
  }.compact
end

def hex2rgb(color=nil)
  color =~ /\A#?(\h\h)(\h\h)(\h\h)\z/i or return
  r = "%.2f" % ($1.hex / 255.0)
  g = "%.2f" % ($2.hex / 255.0)
  b = "%.2f" % ($3.hex / 255.0)
  { red: r, green: g, blue: b }
end

def add_border(area:, top:, bottom:, left:, right:, inner_horizontal:, inner_vertical:)
  reqs = []
  reqs.push(update_borders: {
    range: range(area),
    top: {
      style: top == nil ? "NONE" : top.style,
      color: top == nil ? nil : hex2rgb(top.color)
    },
    bottom: {
      style: bottom == nil ?  "NONE" : bottom.style,
      color: bottom == nil ?  nil : hex2rgb(bottom.color)
    },
    left: {
      style: left == nil ?  "NONE" : left.style,
      color: left == nil ?  nil : hex2rgb(left.color)
    },
    right: {
      style: right == nil ?  "NONE" : right.style,
      color: right == nil ?  nil : hex2rgb(right.color)
    },
    inner_horizontal: {
      style: inner_horizontal == nil ?  "NONE" : inner_horizontal.style,
      color: inner_horizontal == nil ?  nil : hex2rgb(inner_horizontal.color)
    },
    inner_vertical: {
      style: inner_vertical == nil ?  "NONE" : inner_vertical.style,
      color: inner_vertical == nil ?  nil : hex2rgb(inner_vertical.color)
    },
  })
  return reqs
end

def multiple_batch_update(list_of_requests)
  begin
    retries ||= 0
    resp = google.batch_update_spreadsheet($link, { requests: list_of_requests })
    resp
  rescue => e
    retries += 1
    sheet_error(e)
    retry if retries < 3
  end
end