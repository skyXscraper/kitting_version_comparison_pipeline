"""Kitting version comparison pipeline - entry point.

python kvc.py add <folder>              register a folder of result workbooks as a version
python kvc.py truth                     create / update truth/truth.xlsx, then fill True ply / True range in it
python kvc.py score <version>           score one version
python kvc.py compare <v1> <v2> [...]   compare versions on the videos they all ran
python kvc.py rank                      leaderboard of every version
python kvc.py list | remove | config
"""
from kvc.cli import main

if __name__ == "__main__":
    main()
