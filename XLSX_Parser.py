import json
import xlsxwriter

# -------- Load the JSON schedule --------
with open("match_schedule.json", "r", encoding="utf-8") as f:
    final_schedule = json.load(f)

# -------- Helpers to parse entries (supports both old & new JSON shapes) --------
def parse_entry(entry):
    """
    Accepts either:
      ["MS4", ["UCD:", ["Neil"]], ["UCSC:", ["Eric"]]]
    or:
      ["MS4", ["Neil"], ["Eric"]]
    Returns: (event_code, left_label, left_players, right_label, right_players)
    """
    event_code = entry[0]
    # New shape with team label layer
    if len(entry) >= 3 and isinstance(entry[1], list) and len(entry[1]) == 2 and isinstance(entry[1][1], list):
        left_label = entry[1][0].rstrip(":").strip() if entry[1][0] else ""
        left_players = [p.strip() for p in entry[1][1]]
        right_label = entry[2][0].rstrip(":").strip() if entry[2][0] else ""
        right_players = [p.strip() for p in entry[2][1]]
    else:
        # Old shape without team labels
        left_label = ""
        right_label = ""
        left_players = [p.strip() for p in entry[1]]
        right_players = [p.strip() for p in entry[2]]
    return event_code, left_label, left_players, right_label, right_players

def collect_team_labels(schedule_dict):
    """Collect unique team labels (without trailing ':') if present."""
    teams = set()
    for _, courts in schedule_dict.items():
        for _, entry in courts.items():
            try:
                _, llabel, _, rlabel, _ = parse_entry(entry)
                if llabel:
                    teams.add(llabel)
                if rlabel:
                    teams.add(rlabel)
            except Exception:
                pass
    teams = sorted(teams)
    return teams

# -------- Build title + team headers --------
all_teams = collect_team_labels(final_schedule)
title = " vs ".join(all_teams) if all_teams else "Match Schedule"

# Pad to 3 team headers if needed
while len(all_teams) < 3:
    all_teams.append(f"Team {len(all_teams)+1}")

# -------- Excel setup --------
workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column("A:O", 15)
worksheet.set_column("E:E", 12)  # T1 Team
worksheet.set_column("H:H", 12)  # T2 Team
worksheet.set_column("F:G", 20)
worksheet.set_column("K:L", 12)
worksheet.set_column("I:J", 20)
worksheet.set_row(5, 20)
for row in range(7, 201):
    worksheet.set_row(row, 20)

# -------- Formats --------
meet_title_format = workbook.add_format({
    "bold": 1, "underline": 0, "border": 1,
    "align": "center", "valign": "vcenter",
    "fg_color": "#8B5259", "font": "raleway", "size": 22,
})
categories = workbook.add_format({
    "bold": 1, "border": 1, "align": "center", "valign": "vcenter",
    "color": "#FFFFFF", "fg_color": "#355B85", "font": "raleway",
    "size": 12, "bottom": 6,
})
data_blue = workbook.add_format({
    "font": "raleway", "align": "center", "size": 10,
    "border": 1, "fg_color": "#B1D3FA",
    "border_color": "#789FCC",
})
data_blue2 = workbook.add_format({
    "font": "raleway", "align": "center", "size": 10,
    "border": 1, "fg_color": "#A1CAFF",
    "border_color": "#789FCC",
})
data_yellow = workbook.add_format({
    "font": "raleway", "align": "center", "size": 10,
    "border": 1, "fg_color": "#FFDF80",
    "border_color": "#E4C032",
})
data_black = workbook.add_format({
    "font": "raleway", "align": "center", "size": 10,
    "border": 1, "bold": 1, "color": "#FFFFFF", "fg_color": "#333333",
})
data_red = workbook.add_format({
    "font": "raleway", "align": "center", "size": 10,
    "border": 1, "bold": 1, "color": "#FFFFFF", "fg_color": "#79242F",
})
data_grey = workbook.add_format({
    "font": "raleway", "align": "center", "size": 10,
    "border": 1, "fg_color": "#CCCCCC",
})

# -------- Title + Headers --------
worksheet.merge_range("D1:K5", title, meet_title_format)
worksheet.merge_range(
    "A1:C5",
    "Instructions:\n"
    "• Add checkboxes in the 'In progress' column if you like.\n"
    "• Enter 1 in column M if Team 1 (left) wins; enter 1 in column N if Team 2 (right) wins.\n"
    "• Tallies update automatically.",
    data_red
)

worksheet.write('A6', 'Schedule', categories)
worksheet.write('B6', 'Court', categories)
worksheet.write('C6', 'In progress', categories)
worksheet.write('D6', 'Event', categories)
worksheet.write('E6', 'T1 Team', categories)
worksheet.merge_range('F6:G6', 'Team 1', categories)
worksheet.write('H6', 'T2 Team', categories)
worksheet.merge_range('I6:J6', 'Team 2', categories)
worksheet.merge_range('K6:L6', 'Score (Winner First)', categories)
worksheet.merge_range('M6:N6', 'Winner Flags (M=Team1, N=Team2)', categories)

# -------- Row Writing --------
start_counter = 7
for sched_time, courts_dict in final_schedule.items():
    court_count = max(1, len(courts_dict))
    if courts_dict:
        start_cell = "A" + str(start_counter)
        end_cell = "A" + str(start_counter + court_count - 1)
        worksheet.merge_range(f"{start_cell}:{end_cell}", sched_time, categories)

    c_number = 1
    for row in range(start_counter, start_counter + court_count):
        court_cell = "B"+ str(row)
        progress_cell = "C"+ str(row)
        event_cell = "D"+ str(row)
        t1_team_cell = "E" + str(row)
        team1_cell = "F"+ str(row) + ":G" + str(row)
        t2_team_cell = "H" + str(row)
        team2_cell = "I" + str(row) + ":J" + str(row)
        score_cell = "K" + str(row) + ":L" + str(row)
        victory_cell_t1 = "M" + str(row)  # type 1 for left win
        victory_cell_t2 = "N" + str(row)  # type 1 for right win

        entry = courts_dict.get(str(c_number))
        if entry:
            try:
                event_code, left_label, left_players, right_label, right_players = parse_entry(entry)

                left_text = " + ".join(left_players)
                right_text = " + ".join(right_players)

                worksheet.write(progress_cell, "", data_yellow)
                worksheet.write(event_cell, event_code, data_yellow)
                worksheet.write(t1_team_cell, left_label, data_blue2)
                worksheet.write(t2_team_cell, right_label, data_blue2)
                worksheet.merge_range(team1_cell, left_text, data_blue)
                worksheet.merge_range(team2_cell, right_text, data_blue)
                worksheet.merge_range(score_cell, "", data_yellow)

                # Leave winner flags BLANK; user enters 1 in M or N
                worksheet.write(victory_cell_t1, "", data_blue)
                worksheet.write(victory_cell_t2, "", data_blue)

                worksheet.write(court_cell, c_number, data_blue)

            except Exception:
                pass
        c_number += 1
    start_counter += court_count

# -------- Totals / Tally (Dynamic) --------
# Team headers
worksheet.write("L1", all_teams[0], data_red)
worksheet.write("M1", all_teams[1], data_red)
worksheet.write("N1", all_teams[2], data_red)
worksheet.merge_range("L2:N2", "A-Team Tally", data_black)
worksheet.merge_range("L4:N4", "Overall Tally", data_black)

# Build SUMPRODUCT formulas over the full used range
first_row = 7
last_row = start_counter - 1  # last filled row

E_rng = f"$E${first_row}:$E${last_row}"
H_rng = f"$H${first_row}:$H${last_row}"
D_rng = f"$D${first_row}:$D${last_row}"
M_rng = f"$M${first_row}:$M${last_row}"
N_rng = f"$N${first_row}:$N${last_row}"

def overall_formula(header_cell):
    # Sum Team1 wins where left team equals header, plus Team2 wins where right team equals header
    return (
        f"=SUMPRODUCT(({E_rng}={header_cell})*{M_rng})"
        f"+SUMPRODUCT(({H_rng}={header_cell})*{N_rng})"
    )

def ateam_formula(header_cell):
    """
    Restrict to A-Team matches (event codes ending in 1, 2, or 3).
    Excel trick: use ISNUMBER(FIND(...)) instead of VALUE(RIGHT()) for text-safe matching.
    """
    cond_text = (
        f"(ISNUMBER(FIND(\"1\",RIGHT({D_rng},1)))"
        f"+ISNUMBER(FIND(\"2\",RIGHT({D_rng},1)))"
        f"+ISNUMBER(FIND(\"3\",RIGHT({D_rng},1))))"
    )
    return (
        f"=SUMPRODUCT(({E_rng}={header_cell})*{M_rng}*({cond_text}>0))"
        f"+SUMPRODUCT(({H_rng}={header_cell})*{N_rng}*({cond_text}>0))"
    )


# Write formulas for each team column (L/M/N)
worksheet.write_formula("L3", ateam_formula("$L$1"), data_grey)
worksheet.write_formula("M3", ateam_formula("$M$1"), data_grey)
worksheet.write_formula("N3", ateam_formula("$N$1"), data_grey)

worksheet.write_formula("L5", overall_formula("$L$1"), data_grey)
worksheet.write_formula("M5", overall_formula("$M$1"), data_grey)
worksheet.write_formula("N5", overall_formula("$N$1"), data_grey)

workbook.close()
print("✅ Exported result.xlsx — place 1s in M (Team1 wins) or N (Team2 wins); tallies update by team.")
