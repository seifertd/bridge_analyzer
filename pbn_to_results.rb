#!/usr/bin/env ruby

require 'csv'
require_relative 'seat_map'

seats_spec, rest = extract_seats_flag(ARGV)
pbn_file = rest[0]
partner_override = rest[1]

unless pbn_file && seats_spec
  $stderr.puts 'Usage: ruby pbn_to_results.rb --seats <SEATS> <pbn_file> [<partner>]'
  $stderr.puts '  --seats: Doug\'s per-board seats, e.g. "1-5:S,6-10:N,11-20:E,21-25:W"'
  exit 1
end

begin
  SEAT_MAP = parse_seat_map(seats_spec)
rescue ArgumentError => e
  $stderr.puts "Error: #{e.message}"
  exit 1
end

content = File.read(pbn_file)

# Extract target player name from %ACBL_Name header
target_name = "Douglas Seifert"
content.each_line do |line|
  if line =~ /^%ACBL_Name:(.+)/
    target_name = $1.strip.tr('.', ' ')
    break
  end
end

LEFT_OF    = { 'N' => 'E', 'E' => 'S', 'S' => 'W', 'W' => 'N' }.freeze

# Split into board blocks (each starts with [Event)
board_blocks = content.split(/(?=^\[Event\s)/m).select { |b| b.include?('[Board ') }

csv = CSV.new($stdout)
csv << ['Board', 'Dir', 'Contract', 'Score', '% vs Field', '% vs Club',
        '1st Bidder', 'Leader', 'Declarer', 'Vul', 'Lead', 'Bidding', 'Postmortem']

board_blocks.each do |block|
  # Extract PBN tags
  tags = {}
  block.scan(/\[(\w+)\s+"([^"]*)"\]/).each { |k, v| tags[k] = v }

  # Identify seats
  players = { 'N' => tags['North'], 'E' => tags['East'],
              'S' => tags['South'], 'W' => tags['West'] }

  # The PBN normalizes Doug's name into a canonical seat, so use it only to
  # confirm he played this board and to read off his partner's name. His true
  # physical seat (which the file can't encode) comes from the seat map.
  doug_seat = players.find { |_, name| name == target_name }&.first
  next unless doug_seat

  doug_true_seat = seat_for_board(SEAT_MAP, tags['Board'], 'pbn')
  partner_first  = partner_override || players[PARTNER_OF[doug_seat]].split.first
  if !partner_first
    puts "Error: Could not determine partner name from pbn. Provide as 2nd arg"
    exit 2
  end
  dir = %w[E W].include?(doug_true_seat) ? 'EW' : 'NS'

  # Contract: split into characters, N -> NT
  raw_contract = tags['Contract']
  contract = if raw_contract.upcase == 'PASS'
               'PASS'
             else
               raw_contract.chars.map { |c| c == 'N' ? 'NT' : c }.join(' ')
             end

  # Score: the PBN Score tag is from the declarer's perspective.
  # Convert to Doug's perspective: same sign if declarer is on Doug's side,
  # negated if declarer is on the opposing side.
  raw_score     = tags['Score'].to_i
  declarer_dir  = tags['Declarer']
  declarer_side = %w[N S].include?(declarer_dir) ? 'NS' : 'EW'
  doug_score    = (declarer_side == dir) ? raw_score : -raw_score

  # % vs Club: extract the number from ScorePercentage
  club_pct = tags['ScorePercentage'].split.last.to_f

  # % vs Field: look up the NS score in the InstantScoreTable
  ns_score  = %w[N S].include?(declarer_dir) ? raw_score : -raw_score
  field_pct = nil
  in_table  = false
  table_scores = []

  block.each_line do |line|
    if line.include?('[InstantScoreTable')
      in_table = true
      next
    end
    next unless in_table

    stripped = line.strip
    next if stripped.empty?

    if stripped =~ /^(\d+)\s+(-?\d+)\s+([\d.]+)$/
      entry_score = $2.to_i
      entry_pct   = $3.to_f
      table_scores << entry_score
      if entry_score == ns_score
        field_pct = entry_pct
        break
      end
    else
      break
    end
  end

  if field_pct.nil? && !table_scores.empty?
    field_pct = ns_score > table_scores.max ? 100.0 : 0.0
  end

  field_pct = 100.0 - field_pct if field_pct && dir == 'EW'

  # Vulnerable: All -> Both
  vul = tags['Vulnerable']
  vul = case vul
        when 'All'
          'Both'
        when 'None'
          'None'
        else
          vul == dir ? 'Us' : 'Them'
        end

  # Declarer name (relative to Doug's true seat, not the normalized PBN seat)
  declarer_name = if declarer_dir == doug_true_seat
                    'Doug'
                  elsif declarer_dir == PARTNER_OF[doug_true_seat]
                    partner_first
                  else
                    'Defense'
                  end

  # Leader: one seat to the left of declarer
  leader_seat = LEFT_OF[declarer_dir]
  leader_name = if leader_seat == doug_true_seat
                  'Doug'
                elsif leader_seat == PARTNER_OF[doug_true_seat]
                  partner_first
                else
                  'Them'
                end

  csv << [tags['Board'], dir, contract, doug_score,
          format('%.1f', field_pct || 0.0), format('%.1f', club_pct),
          '', leader_name, declarer_name, vul, '', '', '']
end
