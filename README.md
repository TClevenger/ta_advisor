ta_parser.py:  TrustedAdvisor parser.

usage: python3 ta_parser.py FILE

ta_parser.py is a quick and dirty script that parses AWS TrustedAdvisor
reports (Excel format) and removes all sheets where the status is "ok"
or "not_available".  This leaves just the sheets with "warning" or
"error" statuses, making it much easier to warp through the report to
find stuff to actually fix.

This script does no error checking (or even basic checking to see if
it's reading a TrustedAdvisor file), gives no confirmations and
doesn't back up the original, so use at your own risk.

This script requires the openpyxl library (install with 'pip3 install
openpyxl')

