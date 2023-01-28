# FortiEMS/FortiClient Real Time Scan Parser
This script is to automate the review of FortiEMS/FortiClient diagnostic logs. This will take a single .zip file full of .cab files and extract them, pull the real time scan log, and pull any lines where an AV event was detected. Any AV events in the log will be placed in an xlsx workbook associated to the hostname of the device it was found on.

# Usage
1. Python3.10 must be installed on the machine.
2. Ensure any requirements found in the requirements.txt file are installed on your machine.
3. Place the Harvester.py script in an empty directory along with the diag.zip that is downloaded from the FortiEMS portal.
4. Run Harvester.py
5. An .xlsx file will be generated with the results.

NOTE:
  Although fully automated, this script allows manual browsing of the .cab files per associated hostname by extracting the contents into a folder named after the associated hostname.
  In the event that there are issues extracting any .cab files, a directory will be created with the hostname and an appending "Error-".
