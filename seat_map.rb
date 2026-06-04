# Shared parsing for Doug's per-board seat assignments.
#
# Hand-record PBN files normalize the player of interest into a canonical seat
# (N when playing NS, E when playing EW) and a BWS file only records the
# round-1 registration, so neither source reliably tells us which physical
# chair (N vs S, or E vs W) Doug sat in on a given board. That within-pair flip
# only affects the Doug-vs-partner labeling of declarer and opening leader, but
# it cannot be recovered from the files, so the caller must supply it.
#
# The seat spec is a comma-separated list of "<board-or-range>:<DIR>" entries,
# e.g. "1-5:S,6-10:N,11-20:E,21-25:W". DIR is one of N/S/E/W.

SEAT_DIRECTIONS = %w[N S E W].freeze
PARTNER_OF = { 'N' => 'S', 'S' => 'N', 'E' => 'W', 'W' => 'E' }.freeze

# Parse a seat spec into a Hash mapping board number => direction letter.
# Raises ArgumentError on malformed input.
def parse_seat_map(spec)
  if spec.nil? || spec.strip.empty?
    raise ArgumentError, 'missing seat assignments (--seats)'
  end

  map = {}
  spec.split(',').each do |part|
    part = part.strip
    next if part.empty?

    range_str, dir = part.split(':', 2)
    dir = dir.to_s.strip.upcase
    unless SEAT_DIRECTIONS.include?(dir)
      raise ArgumentError, "invalid seat direction in #{part.inspect} (expected N/S/E/W)"
    end

    case range_str.to_s.strip
    when /\A(\d+)-(\d+)\z/
      lo, hi = $1.to_i, $2.to_i
      raise ArgumentError, "invalid board range #{range_str.inspect} (low > high)" if lo > hi
      (lo..hi).each { |b| map[b] = dir }
    when /\A(\d+)\z/
      map[range_str.to_i] = dir
    else
      raise ArgumentError, "invalid board range in #{part.inspect}"
    end
  end

  map
end

# Pull a required "--seats <spec>" / "--seats=<spec>" flag out of argv,
# returning [spec, remaining_args]. Does not parse the spec.
def extract_seats_flag(argv)
  seats = nil
  rest = []
  i = 0
  while i < argv.length
    arg = argv[i]
    if arg == '--seats'
      seats = argv[i + 1]
      i += 2
    elsif arg.start_with?('--seats=')
      seats = arg.split('=', 2)[1]
      i += 1
    else
      rest << arg
      i += 1
    end
  end
  [seats, rest]
end

# Look up Doug's seat for a board, aborting with a clear message if the seat
# map doesn't cover it.
def seat_for_board(seat_map, board, source)
  seat = seat_map[board.to_i]
  unless seat
    abort "Error: no seat assignment for board #{board} in --seats (#{source})"
  end
  seat
end
