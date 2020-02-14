$stdout.sync = true

require 'google/apis/sheets_v4'
require 'googleauth'
require 'googleauth/stores/file_token_store'

APPLICATION_NAME ||= 'Google Sheets Ruby gem'.freeze
SCOPE            ||= Google::Apis::SheetsV4::AUTH_SPREADSHEETS
CREDENTIALS_PATH ||= File.join($APP_ROOT, 'credentials.json').freeze
TOKEN_PATH       ||= File.join($APP_ROOT, 'token.yaml').freeze
OOB_URI          ||= 'urn:ietf:wg:oauth:2.0:oob'.freeze

# Google::Apis.logger = Logger.new(STDERR)
# Google::Apis.logger.level = Logger::DEBUG

helpers do
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
        end
      }
      $google = Google::Apis::SheetsV4::SheetsService.new
      $google.client_options.application_name = APPLICATION_NAME
      $google.authorization = connect.call
      $google
    end
  end

  def sheet_list
    list = google.get_spreadsheet($link).sheets
    list.map do |item|
      vals = item.properties
      {
        id:    vals.sheet_id,
        name:  vals.title,
        color: rgb2hex(vals.tab_color),
      }.compact
    end
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

  def rgb2hex(color=nil)
    color or return
    r = ((color.red   || 0) * 255).to_i
    g = ((color.green || 0) * 255).to_i
    b = ((color.blue  || 0) * 255).to_i
    "#%02x%02x%02x" % [r, g, b]
  end

  def hex2rgb(color=nil)
    color =~ /\A#?(\h\h)(\h\h)(\h\h)\z/i or return
    r = "%.2f" % ($1.hex / 255.0)
    g = "%.2f" % ($2.hex / 255.0)
    b = "%.2f" % ($3.hex / 255.0)
    { red: r, green: g, blue: b }
  end

  def sheet_color(sheet, color=nil) # NOTE: ignores alpha
    reqs = []
    reqs.push(update_sheet_properties: {
      properties: {
        sheet_id: sheet_id(sheet),
        tab_color: hex2rgb(color),
      },
      fields: 'tab_color',
    })
    resp = google.batch_update_spreadsheet($link, { requests: reqs }, {})
  end

  def sheet_format(area, pattern)
    reqs = []
    reqs.push(repeat_cell: {
      range: range(area),
      cell: {
        user_entered_format: {
          number_format: {
            type: "NUMBER",
            pattern: pattern,
          },
        },
      },
      fields: 'user_entered_format.number_format',
    })
    resp = google.batch_update_spreadsheet($link, { requests: reqs }, {})
  end

  def sheet_read(area)
    google.get_spreadsheet_values($link, area).values
  end
  def read(area)
    google.get_spreadsheet_values($link, area).values
  end

  def clear(area)
    google.clear_values($link, area)
  end

  def sheet_save(area, rows)
    gasv = Google::Apis::SheetsV4::ValueRange.new(range: area, values: rows)
    done = google.update_spreadsheet_value($link, area, gasv, value_input_option: "USER_ENTERED")
    done.updated_cells
  end
  def save(area, rows, log=true)
    area.sub!(/^(#\d+)(?=!)/) {|num| sheet_name(num)}
    gasv = Google::Apis::SheetsV4::ValueRange.new(range: area, values: rows)
    done = google.update_spreadsheet_value($link, area, gasv, value_input_option: "USER_ENTERED")
    puts "#{done.updated_cells} cells updated." if log
    rows.count
  end
  def save!(area, rows, log=false)
    save(area, rows, log)
  end

  def update(area, op=nil, **opts, &code)
    rows = read area
    data = rows.map {|cols| cols.map {|cell| touch(cell, op, **opts, &code)}}
    save area, data
  end

  def touch(cell, func=nil, *args, **opts, &code)
    if block_given?
      yield cell
    else
      case func
      when nil  then cell
      when 'jr' then cell + " Junior!"
      else
        if cell.respond_to?(func)
          cell.send(func, *args)
        else
          warn "dude... you gave me the unknown func #{func.inspect}"
          nil
        end
      end
    end
  end

  def touch_cell(cell, func=nil, *args, **opts, &code)
    if block_given?
      yield cell
    else
      case func
      when nil then cell
      when 'to_decimal'
        prec = 2
        if cell[/\A\s*\$?\s*([-+])?\s*\$?\s*([-+])?\s*(\d[,\d]*)?(\.\d*)?\s*\z/]
          sign = "#{$1}#{$2}".squeeze.include?("-") ? "-" : ""
          left = $3.blank? ? "0" : $3.delete(",")
          decs = $4.blank? ? nil : $4
          "%.*f" % [prec, "#{sign}#{left}#{decs}".to_f]
        else
          ""
        end
      when 'to_phone'
        case cell
          when /^1?([2-9]\d\d)(\d{3})(\d{4})$/ then "(#{$1}) #{$2}-#{$3}"
          else ""
        end
      when 'to_yyyymmdd'
        case cell
          when /^((?:19|20)\d{2})(\d{2})(\d{2})$/      then "%s%s%s"       % [$1, $2, $3          ] # YYYYMMDD
          when /^(\d{2})(\d{2})((?:19|20)\d{2})$/      then "%s%s%s"       % [$3, $1, $2          ] # MMDDYYYY
          when /^(\d{1,2})([-\/.])(\d{1,2})\2(\d{4})$/ then "%s%02d%02d"   % [$4, $1.to_i, $3.to_i] # M/D/Y
          when /^(\d{4})([-\/.])(\d{1,2})\2(\d{1,2})$/ then "%s%02d%02d"   % [$1, $3.to_i, $4.to_i] # Y/M/D
          when /^(\d{1,2})([-\/.])(\d{1,2})\2(\d{2})$/
            year = $4.to_i
            year += year < (Time.now.year % 100 + 5) ? 2000 : 1900
            "%04d%02d%02d" % [year, $1.to_i, $3.to_i] # M/D/Y
          else ""
        end
      when 'tune'
        o = {}; opts.each {|e| o[e]=true}
        s = cell
        s = s.downcase.gsub(/\s\s+/, ' ').strip.gsub(/(?<=^| |[\d[:punct:]])([[[:alpha:]]])/i) { $1.upcase } # general case
        s.gsub!(/\b([a-z])\. ?([bcdfghjklmnpqrstvwxyz])\.?(?=\W|$)/i) { "#$1#$2".upcase } # initials (should this be :name only?)
        s.gsub!(/\b([a-z](?:[a-z&&[^aeiouy]]{1,4}))\b/i) { $1.upcase } # uppercase apparent acronyms
        s.gsub!(/\b([djs]r|us|acct|[ai]nn?|apps|ed|erb|esq|grp|in[cj]|of[cf]|st|up)\.?(?=\W|$)/i) { $1.capitalize } # force camel-case
        s.gsub!(/(^|(?<=\d ))?\b(and|at|as|of|the|in|on|or|for|to|by|de l[ao]s?|del?|(el-)|el|las)($)?\b/i) { ($1 || $3 || $4) ? $2.downcase.capitalize : $2.downcase } # prepositions
        s.gsub!(/\b(mc|mac(?=d[ao][a-k,m-z][a-z]|[fgmpw])|[dol]')([a-z])/i) { $1.capitalize + $2.capitalize } # mixed case (Irish)
        s.gsub!(/\b(ahn|an[gh]|al|art[sz]?|ash|e[dnv]|echt|elms|emms|eng|epps|essl|i[mp]|mrs?|ms|ng|ock|o[hm]|ong|orr|orth|ost|ott|oz|sng|tsz|u[br]|ung)\b/i) { $1.capitalize } # if o[:name] # capitalize
        s.gsub!(/(?<=^| |[[:punct:]])(apt?s?|arch|ave?|bldg|blvd|cr?t|co?mn|drv?|elm|end|f[lt]|hts?|ln|old|pkw?y|plc?|prk|pt|r[dm]|spc|s[qt]r?|srt|street|[nesw])\.?(?=\W|$)/i) { $1.capitalize } # if o[:address] # road features
        s.gsub!(/(1st|2nd|3rd|[\d]th|de l[ao]s)\b/i) { $1.downcase } # ordinal numbers
        s.gsub!(/(?<=^|\d |\b[nesw] |\b[ns][ew] )(d?el|las?|los)\b/i) { $1.capitalize } # uppercase (Spanish)
        s.gsub!(/\b(ca|dba|fbo|ihop|mri|ucla|usa|vru|[ns][ew]|i{1,3}v?)\b/i) { $1.upcase } # force uppercase
        s.gsub!(/\b([-@.\w]+\.(?:com|net|io|org))\b/i) { $1.downcase } # domain names, email (a little bastardized...)
        s.gsub!(/# /, '#') # collapse spaces following a number sign
        s.sub!(/[.,#]+$/, '') # nuke any trailing period, comma, or hash signs
        s.sub!(/\bP\.? ?O\.? ?Box/i, 'PO Box') # PO Boxes
        s
      else
        if cell.respond_to?(func)
          cell.send(func, *args)
        else
          warn "dude... you gave me the unknown func #{func.inspect}"
          nil
        end
      end
    end
  end

  def clean_sheet(area)
    rows = sheet_read(area) rescue {}

    todo = Hash[<<~end.scan(/^\s*(.*?)  +(.*?)(?:\s*#.*)?$/)]
      State          upcase
      Source         tune
      Modality       tune
      First Name     tune
      Last Name      tune
      DOB            to_yyyymmdd
      DOA            to_yyyymmdd
      Account Notes  tune
      Law Firm       tune
      Email          downcase
      Phone          to_phone
      Fax            to_phone
      DOS            to_yyyymmdd
      Invoice        to_decimal
      Cost           to_decimal
      Status         tune
      Stage          tune
      Last Contact   to_yyyymmdd
      Contact Notes  tune
    end

    # clean sheet
    seen = 0
    diff = 0
    rows.each_with_index do |cols, r|
      seen += 1
      todo.update(Hash[cols.map.with_index {|name, c| [c, [name, todo[name]]]}]) if seen == 1
      cols.each_with_index do |cell, c|
        name, func = todo[c]
        orig = cell
        cell = touch_cell(cell, func) if func && seen > 1
        if cell != orig
          diff += 1
          cols[c] = cell
        end
      end
    end
    sheet_save(area, rows) if diff > 0
    { summary: "Updated #{diff} cells..." }

    # summarize sheet
    # rows = read area
    # text = rows.map {|cols| cols * "|"} * "\n"
    # file = Tempfile.new(['casava-', '.psv'])
    # file.write text
    # file.flush
    # puts `casava --pipes "#{file.path}"`
  end
end

class String
  def hunks_of(most=200)
    list = []
    if self =~ /^(\w+)!([a-z]+)(\d+):([a-z]+)(\d+)$/i
      base, cone, rone, ctwo, rtwo = $~.captures
      rone = rone.to_i
      rtwo = rtwo.to_i

      head = rone
      if (rtwo - rone) > most
        while head <= rtwo
          tail = [head + (most - 1), rtwo].min
          list.push "#{base}!#{cone}#{head}:#{ctwo}#{tail}"
          head += most
        end
      end
    else
      list.push self
    end
    list
  end
end

# Monthly hearing aid sales by product level and new patient productivity
def report1(area)
  recs = <<~";".sql.show!
    select date_format(pr.purchase_date, "%Y-%m") as `Month`,
           sum(pr.discount_price) as 'Total',
           avg(pr.discount_price) as 'Average',
           avg(if(pl.level=1, pr.discount_price, null)) as '1 Premium',
           avg(if(pl.level=2, pr.discount_price, null)) as '2 High',
           avg(if(pl.level=3, pr.discount_price, null)) as '3 Mid',
           avg(if(pl.level=4, pr.discount_price, null)) as '4 Low',
           avg(if(pl.level=5, pr.discount_price, null)) as '5 Basic',
           sum(if(pl.level=1, pr.discount_price, 0)) / sum(pr.discount_price) as '1 Premium %',
           sum(if(pl.level=2, pr.discount_price, 0)) / sum(pr.discount_price) as '2 High %',
           sum(if(pl.level=3, pr.discount_price, 0)) / sum(pr.discount_price) as '3 Mid %',
           sum(if(pl.level=4, pr.discount_price, 0)) / sum(pr.discount_price) as '4 Low %',
           sum(if(pl.level=5, pr.discount_price, 0)) / sum(pr.discount_price) as '5 Basic %',
           0 as `New Patient Productivity`
    from   purchase      pr
      join hearingaid    ha on ha.hearingaid_id=pr.type_id
      join mhc_levels    pl on pl.id=ha.hearingaid_id left
      join insurance_est ie on ie.transaction_no=pr.transaction_no
    where  pr.purchase_type in ('l_hearing_aid', 'r_hearing_aid')
      and  pr.discount_price>0
      and  pr.archive=0
      and  pr.replaced_purchase_id=0
      and  pr.exchange_purchase_id=0
      and  pr.purchase_date between '2017-01-01' and '2020-12-31'
      and  pl.level>0
      and (ie.insuranceco_id is null or ie.insuranceco_id not in (64,449,450,629))
    group  by `Month`
    order  by `Month`
  ;

  # adjust values
  recs.rows.each_with_index do |cols, i|
    cols.insert 8, nil
  end

  # update sheet
  rows = save area, recs.rows
  sheet_format "#1!K#{rtop = 25}:O#{rtop + rows - 1}", "0%"
  rows
end


# workbook and worksheets
$link = "2I7WROLLau2h7MyqVIJxGzKcR7PiM8_xgc79Q1F_cfwM"
$list = google.get_spreadsheet($link).sheets.map {|s| s.properties }

report1("#1!B25"); save "#1!P2", [[Time.now.strftime("%b %-0d, %Y @ %-0I:%M%P MT")]]

__END__

def basic_filter(sheet_id, count)
  row =  8 # row index (this is the filter header row number)
  col = 21 # col index (this is the filter header right col) 'o' gets 15 ('o'.ord - 'a'.ord + 1)

  batch_update_spreadsheet_request                         = Google::Apis::SheetsV4::BatchUpdateSpreadsheetRequest.new
  batch_update_spreadsheet_request.requests                = []

  clear_basic_filter_request                               = Google::Apis::SheetsV4::ClearBasicFilterRequest.new
  clear_basic_filter_request.sheet_id                      = sheet_id
  batch_update_spreadsheet_request.requests.push Google::Apis::SheetsV4::Request.new(clear_basic_filter: clear_basic_filter_request)

  set_basic_filter_request                                 = Google::Apis::SheetsV4::SetBasicFilterRequest.new
  set_basic_filter_request.filter                          = Google::Apis::SheetsV4::BasicFilter.new
  set_basic_filter_request.filter.range                    = Google::Apis::SheetsV4::GridRange.new
  set_basic_filter_request.filter.range.sheet_id           = sheet_id
  set_basic_filter_request.filter.range.start_row_index    = row - 1
  set_basic_filter_request.filter.range.start_column_index = 1
  set_basic_filter_request.filter.range.end_row_index      = row + count
  set_basic_filter_request.filter.range.end_column_index   = col
  batch_update_spreadsheet_request.requests.push Google::Apis::SheetsV4::Request.new(set_basic_filter: set_basic_filter_request)

  # add_dimension_group_request                   = Google::Apis::SheetsV4::AddDimensionGroupRequest.new
  # add_dimension_group_request.range             = Google::Apis::SheetsV4::DimensionRange.new
  # add_dimension_group_request.range.sheet_id    = 0
  # add_dimension_group_request.range.dimension   = "ROWS"
  # add_dimension_group_request.range.start_index = 2
  # add_dimension_group_request.range.end_index   = 4
  # batch_update_spreadsheet_request.requests.push Google::Apis::SheetsV4::Request.new(add_dimension_group: add_dimension_group_request)

  # add_dimension_group_request                   = Google::Apis::SheetsV4::AddDimensionGroupRequest.new
  # add_dimension_group_request.range             = Google::Apis::SheetsV4::DimensionRange.new
  # add_dimension_group_request.range.sheet_id    = 0
  # add_dimension_group_request.range.dimension   = "ROWS"
  # add_dimension_group_request.range.start_index = 13
  # add_dimension_group_request.range.end_index   = 17
  # batch_update_spreadsheet_request.requests.push Google::Apis::SheetsV4::Request.new(add_dimension_group: add_dimension_group_request)

  google.batch_update_spreadsheet($url, batch_update_spreadsheet_request)
end

area = 'Report!B9:K'

count = report area
basic_filter $list[0].sheet_id, count
