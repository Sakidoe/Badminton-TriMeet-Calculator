

# Badminton Trimeet Calculator

This is a backend program that allows 3 teams to input rosters through .xlsx files. It then translates the data into .json objects, parses it into a greedy algorithm to produce a schedule, then returns an editable and scripted .xlsx file to be used.

Steps:
1. Import 3 .xlsx files with each team's respective rosters, in the format of ![image](https://github.com/Sakidoe/Badminton_Meet_Calculator/assets/114327608/a7b267f5-6bc4-4611-94c2-c8b3d9a95b0e)
Feel free to utilize the template .xlsx files, and also please double check all teams have equal roster sizes.
2. Python3 install all packages that all 4 programs use.
3. Run
```
python3 xlsx-json.py
```
and run through the UI to fully complete adding in rosters, then make the json files of the teams.
 4. Run
```
python3 meet-scheduler.py
```
5.  Run
```
python3 XLSX_Parser.py
```
