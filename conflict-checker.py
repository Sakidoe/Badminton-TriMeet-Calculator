import json
import pandas as pd
from collections import defaultdict
from pathlib import Path

def flatten_players(entry):
    """Extract all player names from a match entry"""
    players = []
    for side in entry[1:]:
        if isinstance(side, list) and len(side) == 2 and isinstance(side[1], list):
            players.extend(side[1])
    return [p.strip() for p in players if p.strip()]

def check_conflicts(schedule):
    """Find conflicts (same player playing multiple matches in same slot)"""
    conflicts = {}
    for timeslot, courts in schedule.items():
        player_to_matches = defaultdict(list)
        for court, match in courts.items():
            match_id = match[0]
            players = flatten_players(match)
            for p in players:
                player_to_matches[p].append(match_id)
        # Conflicts = players appearing >1 time
        conflicts[timeslot] = {p: m for p, m in player_to_matches.items() if len(m) > 1}
    return {t: c for t, c in conflicts.items() if c}

def check_back_to_back(schedule):
    """Find players who have matches in consecutive timeslots"""
    back_to_back = defaultdict(list)
    all_times = sorted(schedule.keys(), key=lambda t: (
        int(t.split(':')[0]) * 60 + int(t.split(':')[1])
    ))
    prev_slot_players = set()

    for t in all_times:
        current_players = set()
        for match in schedule[t].values():
            current_players.update(flatten_players(match))
        # Find intersection
        overlap = current_players & prev_slot_players
        for p in overlap:
            back_to_back[p].append((prev_slot, t))
        prev_slot_players = current_players
        prev_slot = t

    return dict(back_to_back)

def player_summary(schedule):
    """Return match counts and first/last appearance for each player"""
    summary = defaultdict(lambda: {"matches": 0, "slots": []})
    for t, matches in schedule.items():
        for m in matches.values():
            for p in flatten_players(m):
                summary[p]["matches"] += 1
                summary[p]["slots"].append(t)
    for p, info in summary.items():
        info["slots"].sort(key=lambda x: int(x.split(':')[0]) * 60 + int(x.split(':')[1]))
        info["first"] = info["slots"][0]
        info["last"] = info["slots"][-1]
    return summary

def save_to_excel(conflicts, back_to_back, summary, output="conflict_report.xlsx"):
    with pd.ExcelWriter(output) as writer:
        # Conflicts
        if conflicts:
            df_conflicts = pd.DataFrame([
                {"Time": t, "Player": p, "Matches": ", ".join(m)}
                for t, data in conflicts.items() for p, m in data.items()
            ])
            df_conflicts.to_excel(writer, index=False, sheet_name="Conflicts")
        else:
            pd.DataFrame([{"Message": "No conflicts found"}]).to_excel(writer, index=False, sheet_name="Conflicts")

        # Back-to-back
        if back_to_back:
            df_b2b = pd.DataFrame([
                {"Player": p, "Consecutive Slots": " → ".join([a + " - " + b for a, b in slots])}
                for p, slots in back_to_back.items()
            ])
            df_b2b.to_excel(writer, index=False, sheet_name="BackToBack")
        else:
            pd.DataFrame([{"Message": "No back-to-back matches found"}]).to_excel(writer, index=False, sheet_name="BackToBack")

        # Summary
        df_summary = pd.DataFrame([
            {"Player": p, "Total Matches": v["matches"], "First Slot": v["first"], "Last Slot": v["last"]}
            for p, v in summary.items()
        ])
        df_summary.sort_values(by="Total Matches", ascending=False).to_excel(writer, index=False, sheet_name="Summary")

    print(f"✅ Conflict report saved to {output}")

def main():
    # Load schedule JSON
    path = Path("match_schedule.json")
    if not path.exists():
        print("❌ Missing match_schedule.json file.")
        return

    with open(path, "r", encoding="utf-8") as f:
        schedule = json.load(f)

    conflicts = check_conflicts(schedule)
    back_to_back = check_back_to_back(schedule)
    summary = player_summary(schedule)

    save_to_excel(conflicts, back_to_back, summary)

if __name__ == "__main__":
    main()
