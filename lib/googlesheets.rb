# ============================================================================
# googlesheets - Ruby gem for Google Sheets
#
# Author: Steve Shreeve (steve.shreeve@gmail.com)
#   Date: March 16, 2023
# ============================================================================

# See https://googleapis.dev/ruby/google-api-client/latest/Google/Apis/SheetsV4/Request.html

require "google/apis/sheets_v4"
require "googleauth"
require "googleauth/stores/file_token_store"

class Object
  def blank?
    respond_to?(:empty?) or return !self
    empty? or respond_to?(:strip) && strip.empty?
  end unless defined? blank?
end

class GoogleSheets
  VERSION = "0.7.0"

  attr_accessor :api

  def initialize(ssid, **opts)
    @ssid = ssid =~ /^https?:/ ? ssid.split("/")[5] : ssid # spreadsheet id
    @wsid = ssid =~ /(?<=#gid=)(\d+)(?:!([a-z\d:]+))?$/i ? $1.to_i : "#1" # worksheet id
    @rect = $2 ? $2.sub(/(?<=\D):/, "1:") : "A:ZZ"

    @json = opts[:credentials] || "credentials.json"
    @yaml = opts[:token      ] || "token.yaml"

    if !File.exist?(@json)
      base = File.expand_path(File.join(ENV["HOME"], ".google"))
      abort "unable to file Google API credentials" unless File.exist?(base)
      @json = File.join(base, @json)
      @yaml = File.join(base, @yaml)
    end

    if opts[:debug] == true
      $stdout.sync = true
      Google::Apis.logger = Logger.new(STDERR)
      Google::Apis.logger.level = Logger::DEBUG
    end

    @api = Google::Apis::SheetsV4::SheetsService.new
    @api.client_options.application_name    = opts[:app] || "Ruby"
    @api.client_options.application_version = opts[:ver] || "1.0.0"
    @api.authorization = authorize
  end

  def authorize
    idno = Google::Auth::ClientId.from_file(@json)
    repo = Google::Auth::Stores::FileTokenStore.new(file: @yaml)
    auth = Google::Auth::UserAuthorizer.new(idno, Google::Apis::SheetsV4::AUTH_SPREADSHEETS, repo)
    oobs = "urn:ietf:wg:oauth:2.0:oob"
    user = "default"
    info = auth.get_credentials(user) || begin
      href = auth.get_authorization_url(base_url: oobs)
      puts "Open the following URL and paste the code here:\n" + href
      info = auth.get_and_store_credentials_from_code(user_id: user, code: gets, base_url: oobs)
    end
  end

  def biject(x) # a=1, z=26, aa=27, az=52, ba=53, aaa=703
    case x
    when String
      x.each_char.inject(0) {|n,c| (n * 26) + (c.ord & 31) }
    when Integer
      s = []
      s << (((x -= 1) % 26) + 65).chr && x /= 26 while x > 0
      s.reverse.join
    end
  end

  def hex2rgb(color=nil)
    color =~ /\A#?(?:(\h\h)(\h\h)(\h\h)|(\h)(\h)(\h))\z/ or return
    r, g, b = $1 ? [$1, $2, $3] : [$4*2, $5*2, $6*2]
    r = "%.2f" % (r.hex / 255.0)
    g = "%.2f" % (g.hex / 255.0)
    b = "%.2f" % (b.hex / 255.0)
    { red: r, green: g, blue: b }
  end

  def rgb2hex(color=nil)
    color or return
    r = ((color.red   || 0) * 255).to_i
    g = ((color.green || 0) * 255).to_i
    b = ((color.blue  || 0) * 255).to_i
    "#%02x%02x%02x" % [r, g, b]
  end

  def filter_criteria(hash)
    hash.inject({}) do |h, (k,v)|
      l = Array(v)
      h[biject(k.to_s) - 1] = {
        condition: {
          type: "TEXT_EQ",
          values: l.map {|e| { user_entered_value: e} },
        }
      }
      h
    end
  end

  def range(area)
    sh, rc = area.split("!", 2); rc, sh = sh, nil if sh.nil?
    as, ae = rc.split(":", 2); ae ||= as
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

  # --------------------------------------------------------------------------

  def sheets
    @sheets ||= api.get_spreadsheet(@ssid).sheets.map {|item| item.properties }
  end

  def sheets!
    @sheets = nil
    sheets
  end

  def resolve_area(area)
    if area.blank?
      wsid = sheet_name
    elsif area =~ /!/
      wsid = sheet_name($`)
      rect = $'
    elsif area =~ /:/ or area =~ /^[a-z]{1,3}\d+$/i
      wsid = sheet_name
      rect = area
    else
      wsid = sheet_name(area)
    end
    rect = rect ? rect.sub(/(?<=\D):/, "1:") : @rect
    "#{wsid}!#{rect}"
  end

  def sheet_name(obj=nil)
    case obj ||= @wsid
      when /^#(\d+)(?:!?|$)/ then sheets[$1.to_i - 1].title
      when "", 0, nil        then sheets.first.title
      when Integer           then sheets.find {|item| item.sheet_id == obj}&.title
      else obj
    end
  end

  def sheet_clear(area)
    area = resolve_area(area)
    api.clear_values(@ssid, area)
  end

  def sheet_color(curr, color=nil) # NOTE: ignores alpha
    reqs = []
    reqs.push(update_sheet_properties: {
      properties: {
        sheet_id: sheet_id(curr),
        tab_color: hex2rgb(color),
      },
      fields: "tab_color",
    })
    resp = api.batch_update_spreadsheet(@ssid, { requests: reqs }, {})
    true
  end

  def sheet_filter(area, filt=nil)
    area = resolve_area(area)
    range = range(area)
    criteria = filter_criteria(filt) if filt
    reqs = []
    reqs.push(clear_basic_filter: { sheet_id: range[:sheet_id] })
    reqs.push(set_basic_filter: { filter: { range: range, criteria: criteria}.compact })
    resp = api.batch_update_spreadsheet(@ssid, { requests: reqs }, {})
    true
  end

  def sheet_format(area, pattern)
    area = resolve_area(area)
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
      fields: "user_entered_format.number_format",
    })
    resp = api.batch_update_spreadsheet(@ssid, { requests: reqs }, {})
    true
  end

  def sheet_id(obj)
    case obj
      when /^#(\d+)(?:!?|$)/ then sheets[$1.to_i - 1].sheet_id
      when Integer           then obj
      else sheets.find {|item| item.title == obj}&.sheet_id
    end
  end

  def sheet_list
    sheets.map do |item|
      {
        id:    item.sheet_id,
        name:  item.title,
        color: rgb2hex(item.tab_color),
      }.compact
    end
  end

  def sheet_rename(curr, name=nil)
    shid = sheet_id(curr)
    name ||= yield(sheet_name(shid)) if block_given?

    reqs = []
    reqs.push(update_sheet_properties: {
      properties: {
        sheet_id: shid,
        title: name,
      },
      fields: "title",
    })
    resp = api.batch_update_spreadsheet(@ssid, { requests: reqs }, {})
    true
  end

  def sheet_read(area=nil)
    area = resolve_area(area)
    api.get_spreadsheet_values(@ssid, area).values
  end

  def sheet_save(area, rows, log=false)
    area = resolve_area(area)
    gasv = Google::Apis::SheetsV4::ValueRange.new(range: area, values: rows)
    done = api.update_spreadsheet_value(@ssid, area, gasv, value_input_option: "USER_ENTERED")
    puts "#{done.updated_cells} cells updated." if log
    done.updated_cells
  end

  def sheet_save!(area, rows)
    sheet_save(area, rows, true)
  end
end
