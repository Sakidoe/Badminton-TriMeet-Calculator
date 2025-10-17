import json
from collections import defaultdict
from datetime import datetime, timedelta

# Prioritize WD so it doesn't get squeezed out, then WS, MD, XD, MS
EVENTS = ["WD", "WS", "MD", "XD", "MS"]

def build_matches(meet, teams):
    """Return list of match dicts for tri-meet A-B, A-C, B-C at matching ranks."""
    def team_rank_players(team, event, rank_key):
        return meet[team][event].get(rank_key, {}).get("Player Name")

    pairings = [(teams[0], teams[1]), (teams[0], teams[2]), (teams[1], teams[2])]
    matches = []
    for event in EVENTS:
        # collect all ranks that appear for this event across the three teams
        all_ranks = set()
        for t in teams:
            all_ranks |= set(meet[t][event].keys())
        # for each rank, create the two-school matches that exist
        def rank_key_fn(r):
            try:
                return int(r.split()[-1])
            except Exception:
                return 10**9
        for rank in sorted(all_ranks, key=rank_key_fn):
            for t1, t2 in pairings:
                p1 = team_rank_players(t1, event, rank)
                p2 = team_rank_players(t2, event, rank)
                if p1 and p2:
                    matches.append({
                        "event": event,
                        "rank": rank,  # e.g., "Rank 3"
                        "teams": (t1, t2),
                        "players": {
                            t1: [f"{t1}:{name}" for name in p1],
                            t2: [f"{t2}:{name}" for name in p2],
                        }
                    })
    return matches

def conflict_sets(matches):
    """Build conflict sets: two matches conflict if they share any player."""
    player_to_matches = defaultdict(list)
    for i, m in enumerate(matches):
        for players in m["players"].values():
            for p in players:
                player_to_matches[p].append(i)

    conflicts = [set() for _ in matches]
    for lst in player_to_matches.values():
        for i in lst:
            for j in lst:
                if i != j:
                    conflicts[i].add(j)
    return conflicts

def build_slot_times(slot_minutes=20, windows=None):
    """
    windows: list of tuples [(start_str, end_str), ...], e.g. [("4:30","12:00"), ("13:00","19:00")]
    Returns a list of (slot_start_dt, slot_end_dt) with step=slot_minutes.
    """
    if windows is None:
        windows = [("10:20", "12:00"), ("13:00", "19:00")]
    slots = []
    for start_s, end_s in windows:
        start = datetime.strptime(start_s, "%H:%M")
        end = datetime.strptime(end_s, "%H:%M")
        t = start
        while t + timedelta(minutes=slot_minutes) <= end:
            slots.append((t, t + timedelta(minutes=slot_minutes)))
            t += timedelta(minutes=slot_minutes)
    return slots

def schedule_matches(matches, courts=6, max_slots=None):
    """
    Greedy graph coloring with per-slot capacity and a fixed number of slots.
    Returns:
      match_slot       : list of slot indices for each match (or -1 if unscheduled)
      slot_matches     : list (len=max_slots) of lists of match indices scheduled in that slot
      unscheduled_list : list of match indices that couldn't be placed
    """
    conflicts = conflict_sets(matches)
    # order by degree (most constrained first)
    order = sorted(range(len(matches)), key=lambda i: len(conflicts[i]), reverse=True)

    if max_slots is None:
        raise ValueError("max_slots must be provided (use build_slot_times to define windows).")

    slot_matches = [[] for _ in range(max_slots)]
    match_slot = [-1] * len(matches)
    unscheduled = []

    for mid in order:
        placed = False
        for s in range(max_slots):
            if len(slot_matches[s]) >= courts:
                continue
            # check for any player conflict in this slot
            if any(other in conflicts[mid] for other in slot_matches[s]):
                continue
            # ok to place
            slot_matches[s].append(mid)
            match_slot[mid] = s
            placed = True
            break
        if not placed:
            unscheduled.append(mid)

    return match_slot, slot_matches, unscheduled

def assign_courts(slot_matches):
    """Within each slot, assign courts 1..N (as integers)."""
    result = {}
    for s, mids in enumerate(slot_matches):
        result[s] = [(i+1, mid) for i, mid in enumerate(mids)]
    return result

def _rank_number(rank_str):
    """Convert 'Rank 7' -> 7; safe fallback to 0."""
    try:
        return int(rank_str.split()[-1])
    except Exception:
        return 0

def _fmt_time_key_12h(dt_obj):
    """
    Format like: 11:00, 11:30, 12:00, 12:30, 01:00, 01:30, ...
    (12-hour clock, zero-padded, no AM/PM)
    """
    h = dt_obj.hour % 12 or 12
    return f"{h:02d}:{dt_obj.strftime('%M')}"

def schedule_to_json(matches, slot_matches, slot_times):
    """
    Build JSON in the requested format with an extra team label layer:
      "6": ["MS4", ["UCD:", ["Neil Patel"]], ["UCSC:", ["Eric Wang"]]]
    """
    # Pre-init all time keys (including empty ones)
    out = { _fmt_time_key_12h(start): {} for (start, _end) in slot_times }

    courts_by_slot = assign_courts(slot_matches)

    for s, (slot_start, _slot_end) in enumerate(slot_times):
        time_key = _fmt_time_key_12h(slot_start)
        for court_num, mid in courts_by_slot.get(s, []):
            m = matches[mid]
            ev = m["event"]                      # "MD", "MS", ...
            rnum = _rank_number(m["rank"])       # int
            code = f"{ev}{rnum}"                 # e.g., "MD3"

            t1, t2 = m["teams"]                  # team codes like "UCD", "UCSC"
            # Strip "TEAM:" prefixes and trim whitespace on names
            p1 = [name.split(":", 1)[1].strip() for name in m["players"][t1]]
            p2 = [name.split(":", 1)[1].strip() for name in m["players"][t2]]

            # Build the nested structure with team labels
            left  = [f"{t1}:", p1]
            right = [f"{t2}:", p2]

            # Courts must be string keys: "1", "2", ...
            out[time_key][str(court_num)] = [code, left, right]

    return out

def make_schedule(meet_json_path, teams, courts=6, slot_minutes=20,
                  windows=(("10:20","12:00"), ("13:00","19:00"))):
    with open(meet_json_path, "r", encoding="utf-8") as f:
        meet = json.load(f)

    # ensure all event keys exist for all teams
    for t in teams:
        for e in EVENTS:
            meet[t].setdefault(e, {})

    # Build all matches (now includes WD) with priority ordering
    matches = build_matches(meet, teams)

    # Build custom slot times for the day
    slot_times = build_slot_times(slot_minutes=slot_minutes, windows=windows)
    max_slots = len(slot_times)

    # Schedule with fixed windows and courts
    match_slot, slot_matches, unscheduled = schedule_matches(
        matches, courts=courts, max_slots=max_slots
    )

    # Build JSON object in requested structure
    schedule_json = schedule_to_json(matches, slot_matches, slot_times)

    # Prepare a simple summary
    summary = {
        "total_matches": len(matches),
        "scheduled_matches": len(matches) - len(unscheduled),
        "unscheduled_matches": len(unscheduled),
    }

    # If anything didn't fit, list a short warning + the first few unscheduled
    warning = ""
    if unscheduled:
        preview = []
        for mid in unscheduled[:10]:
            m = matches[mid]
            preview.append(f"{m['event']} {m['rank']} — {m['teams'][0]} vs {m['teams'][1]}")
        warning = (
            f"WARNING: {len(unscheduled)} matches could not be scheduled within the windows. "
            f"Consider adding slots or reordering priorities.\n"
            + "\n".join(f"  - {line}" for line in preview)
        )

    return matches, match_slot, slot_matches, schedule_json, summary, warning

# ------------------- example usage -------------------
if __name__ == "__main__":
    teams = ["UCSC", "UCD", "SJSU"]
    _, _, _, schedule_json, summary, warning = make_schedule(
        "save.json",
        teams,
        courts=6,
        slot_minutes=20,
        windows=(("10:20","12:00"), ("13:00","19:00"))  # 10:20–12, 13–19 in 20-min slots
    )

    # Write the JSON schedule out in the requested format
    with open("match_schedule.json", "w", encoding="utf-8") as f:
        json.dump(schedule_json, f, indent=3, ensure_ascii=False)

    print("--- Summary ---")
    print(summary)
    if warning:
        print(warning)
    else:
        print("All matches scheduled within the given windows.")
    print("Saved JSON schedule to match_schedule.json")
