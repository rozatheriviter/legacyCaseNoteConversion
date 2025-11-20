#!/bin/bash

# --- GPL Compliance Notice ---
cat << 'EOF'
Legacy Case Note Conversion Tool Chain
Copyright (C) 2025  Roza Nil

This program comes with ABSOLUTELY NO WARRANTY; for details see the
GNU General Public License.
This is free software, and you are welcome to redistribute it
under certain conditions; see <https://www.gnu.org/licenses/> for details.

EOF
# -----------------------------

python -m venv venv
source venv/bin/activate

echo "---starting process---"
echo "step 1: running extracting script"
python3 extract.py

echo "running data filtering and conversion to csv script."
python3 convertDOCtoCSV.py
python3 csv2xlsx.py

echo "script complete"
exit