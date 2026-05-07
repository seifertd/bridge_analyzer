require 'csv'

# Command line arguments
if ARGV.size < 1 || ARGV.size > 2
  puts "Usage: ruby export_to_sheet.rb <BWSFILE> [PARTNER_NAME]"
  puts "  BWSFILE: Path to the .BWS file"
  puts "  PARTNER_NAME: (optional) Your partner's name - overrides BWS file if provided"
  exit 1
end

MDBEXPORT = "mdb-export"
DATAFILE = ARGV[0]
PARTNER_NAME_OVERRIDE = ARGV[1]  # May be nil

DEALER = ['N', 'E', 'S', 'W']
VUL = ['O', 'NS', 'EW', 'B', 'NS', 'EW', 'B', 'O', 'EW', 'B', 'O', 'NS', 'B', 'O', 'NS', 'EW']
TRICK_VAL = {"NT" => 30, "S" => 30, "H" => 30, "C" => 20, "D" => 20}

CSV::Converters[:my_dt] = lambda do |s|
  begin
    if s.include?('00:00:00')
      Date.strptime(s, '%m/%d/%y')
    elsif s.include?('12/30/99')
      Time.strptime(s.sub('12/30/99 ', ''), '%H:%M:%S')
    else
      Time.strptime(s, '%m/%d/%y %H:%M:%S')
    end
  rescue Exception => e
    s
  end
end

# Load player data first to find Doug
players_csv = `#{MDBEXPORT} -B #{DATAFILE} PlayerNumbers`
players = CSV.parse(players_csv, headers: true, converters: [:numeric, :date_time])

# Find Doug's entry (search for "Seifert" in name)
doug_entry = players.find { |p| p['Name'] =~ /Seifert/i }

unless doug_entry
  puts "Error: Could not find 'Seifert' in PlayerNumbers table"
  puts "Available players:"
  players.each { |p| puts "  Table #{p['Table']}, #{p['Direction']}: #{p['Name']}" }
  exit 1
end

MY_TABLE = doug_entry['Table'].to_i        # physical starting table
MY_SECTION = doug_entry['Section'].to_i
MY_DIRECTION = doug_entry['Direction']

# Derive pair number from RoundData (physical table != pair number)
round_data = CSV.parse(`#{MDBEXPORT} -B #{DATAFILE} RoundData`, headers: true, converters: [:numeric])
my_round1 = round_data.find { |r| r['Section'] == MY_SECTION && r['Table'] == MY_TABLE && r['Round'] == 1 }
MY_PAIR = ['N', 'S'].include?(MY_DIRECTION) ? my_round1['NSPair'] : my_round1['EWPair']

# Determine partner direction (opposite direction in partnership)
PARTNER_DIRECTION = case MY_DIRECTION
when 'N' then 'S'
when 'S' then 'N'
when 'E' then 'W'
when 'W' then 'E'
end

# Find partner's entry
partner_entry = players.find { |p| p['Table'].to_i == MY_TABLE && p['Direction'] == PARTNER_DIRECTION }

# Determine partner name
PARTNER_NAME = if PARTNER_NAME_OVERRIDE
  # Command line override takes precedence
  PARTNER_NAME_OVERRIDE
elsif partner_entry && partner_entry['Name'] && !partner_entry['Name'].to_s.strip.empty?
  # Extract first name from BWS file (e.g., "Carl Ebeling" -> "Carl")
  partner_entry['Name'].strip.split(/\s+/).first
else
  # Partner name missing from BWS and no override provided
  puts "Error: Partner name is blank in the BWS file for Table #{MY_TABLE}, Direction #{PARTNER_DIRECTION}"
  puts "Please provide partner name on command line:"
  puts "  ruby export_to_sheet.rb #{DATAFILE} \"Partner Name\""
  exit 1
end

# Build players lookup hash
players_hash = players.inject({}) { |h, p| h["#{p['Table']}-#{p['Direction']}"] = p; h }

# Load all data
data = CSV.parse(`#{MDBEXPORT} -B #{DATAFILE} ReceivedData`,
                 headers: true, converters: [:numeric, :my_dt])

# Calculate bridge scoring
def score(contract, result, vul)
  return 0 if contract.nil? || contract.strip.empty? || contract.strip.upcase == 'PASS'
  book, suit, doubling = contract.split(" ")
  book = book.to_i
  return 0 if result.nil?
  result = 0 if result == '='
  
  # Doubling multiplier: undoubled=1, doubled=2, redoubled=4
  dbl = doubling.nil? ? 1 : (doubling == 'x' ? 2 : 4)
  
  if result >= 0  # Made contract
    # Trick score
    trick_score = book * TRICK_VAL[suit] * dbl
    trick_score += 10 * dbl if suit == 'NT'  # NT first trick bonus
    
    # Overtricks
    overtrick_value = TRICK_VAL[suit]
    overtrick_value = (vul ? 100 : 50) * dbl if dbl > 1  # Doubled overtricks worth more
    overtrick_score = overtrick_value * result
    
    # Game bonus: 300/500 for game, 50 for partscore
    game_bonus = trick_score >= 100 ? (vul ? 500 : 300) : 50
    game_bonus += 25 * dbl if dbl > 1  # Bonus for making doubled/redoubled
    
    # Slam bonus
    slam_bonus = if book == 6
      vul ? 750 : 500  # Small slam
    elsif book == 7
      vul ? 1500 : 1000  # Grand slam
    else
      0
    end
    
    trick_score + overtrick_score + game_bonus + slam_bonus
  else  # Down
    # Undertrick penalties (negative result = number of tricks down)
    if vul
      case dbl
      when 1 then 100 * result  # -100 per trick
      when 2 then -200 + 300 * (result + 1)  # First down -200, rest -300
      else -400 + 600 * (result + 1)  # Redoubled: first -400, rest -600
      end
    else  # Not vulnerable
      case dbl
      when 1 then 50 * result  # -50 per trick
      when 2
        # First down -100, 2nd/3rd -200 each, 4th+ -300 each
        first_tricks = [result + 1, -2].max  # -1 or -2
        extra_tricks = [result + 3, 0].min   # -3 or more
        -100 + 200 * first_tricks + 300 * extra_tricks
      else  # Redoubled
        first_tricks = [result + 1, -2].max
        extra_tricks = [result + 3, 0].min
        -200 + 400 * first_tricks + 600 * extra_tricks
      end
    end
  end
end

def declarer_vul(vul, declarer)
  vul == 'B' || vul.include?(declarer)
end

def opening_leader(declarer)
  # Leader is left of declarer (clockwise)
  {'N' => 'E', 'E' => 'S', 'S' => 'W', 'W' => 'N'}[declarer]
end

# When playing EW, the N player sits E and S player sits W (standard compass rotation)
ROTATE = {'N' => 'E', 'E' => 'S', 'S' => 'W', 'W' => 'N'}
MY_EW_DIRECTION      = ROTATE[MY_DIRECTION]
PARTNER_EW_DIRECTION = ROTATE[PARTNER_DIRECTION]

def position_to_readable(position, my_compass, partner_compass)
  case position
  when my_compass      then 'Doug'
  when partner_compass then PARTNER_NAME
  else 'Them'
  end
end

# Calculate MP scores for each board
def calc_mp_scores(data, board)
  results = data.find_all { |r| r['Board'] == board }.map do |r|
    vul = VUL[(board - 1) % VUL.size]
    dec = r['NS/EW']
    dv = declarer_vul(vul, dec)
    s = score(r['Contract'], r['Result'], dv)
    s = ['N', 'S'].include?(dec) ? s : -s
    [r, s]
  end
  
  num_results = results.size - 1
  return {} if num_results < 0
  
  scores = results.map { |r, s| s }
  
  mp_by_row = {}
  results.each do |r, s|
    eq_adj = 0.5 * (scores.count(s) - 1)
    gr_adj = scores.count { |s2| s2 > s }
    mp_score = num_results - eq_adj - gr_adj
    mp_per = (mp_score / num_results * 100.0).round(1)
    mp_by_row[r] = {score: mp_score, percent: mp_per}
  end
  
  mp_by_row
end

# Get my boards by pair number
my_boards = data.find_all { |r| r['PairNS'] == MY_PAIR || r['PairEW'] == MY_PAIR }

# Build output CSV
output = []
boards_played = my_boards.map { |r| r['Board'] }.sort.uniq

boards_played.each do |board_num|
  mp_data = calc_mp_scores(data, board_num)
  my_result = my_boards.find { |r| r['Board'] == board_num }
  
  # Determine my direction for this board
  my_dir = my_result['PairNS'] == MY_PAIR ? 'NS' : 'EW'
  
  # Calculate score from my perspective
  vul = VUL[(board_num - 1) % VUL.size]
  dec_dir = my_result['NS/EW']
  dec_vul = declarer_vul(vul, dec_dir)
  raw_score = score(my_result['Contract'], my_result['Result'], dec_vul)
  my_score = ['N', 'S'].include?(dec_dir) ? raw_score : -raw_score
  my_score = -my_score if my_dir == 'EW'
  
  passed_out = my_result['Contract'].nil? || my_result['Contract'].to_s.strip.empty? ||
               my_result['Contract'].to_s.strip.upcase == 'PASS'
  dealer = DEALER[(board_num - 1) % DEALER.size]
  mp_info = mp_data[my_result]
  mp_pct = my_dir == 'EW' ? (100.0 - mp_info[:percent]).round(1) : mp_info[:percent]

  # Human-readable vulnerability
  vul_readable = case vul
  when 'O' then 'None'
  when 'B' then 'Both'
  when 'NS' then my_dir == 'NS' ? 'Us' : 'Them'
  when 'EW' then my_dir == 'EW' ? 'Us' : 'Them'
  end

  my_compass      = my_dir == 'NS' ? MY_DIRECTION      : MY_EW_DIRECTION
  partner_compass = my_dir == 'NS' ? PARTNER_DIRECTION : PARTNER_EW_DIRECTION

  if passed_out
    leader_readable = ''
    declarer_readable = ''
  else
    leader = opening_leader(dec_dir)
    leader_readable = position_to_readable(leader, my_compass, partner_compass)
    dec_readable = position_to_readable(dec_dir, my_compass, partner_compass)
    declarer_readable = dec_readable == 'Them' ? 'Defense' : dec_readable
  end

  output << {
    'Board' => board_num,
    'Dir' => my_dir,
    'Contract' => passed_out ? '' : my_result['Contract'],
    'Score' => my_score,
    '% vs Field' => mp_pct,
    '% vs Club' => mp_pct,
    '1st Bidder' => '',
    'Leader' => leader_readable,
    'Declarer' => declarer_readable,
    'Vul' => vul_readable,
    'Lead' => '',
    'Bidding' => '',
    'Postmortem' => ''
  }
end

# Write CSV
CSV($stdout) do |csv|
  csv << output.first.keys
  output.each { |row| csv << row.values }
end
