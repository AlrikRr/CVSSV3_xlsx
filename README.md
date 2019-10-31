# CVSS v3 Serpico Output
This excel file has a JSON output for each Findings. This JSON output is compatible with the import in the base SERPICO Project.
You can easily with a script (PowerShell / Python) retrieve each JSON output and insert it into a JSON file for import.
CF: https://www.github.com/SerpicoProject


# CVSS v3 Brute
This file is here as an example to understand the formulas used to calculate the CVSS score.
You can use this file to resume the calculations and thus make your own version.

# CVSS v3 Line
This file is the default file with a VECTOR output. You can copy the first line so apply it to as many "cases" as you want. This file has no JSON output

# Powershell_export_json.ps1
This PowerShell script will retrieve the JSS cells in the CVSS v3 SERPICO file to generate a JSON file ready for import into SERPICO. Remember to look at the code of the script before executing it. There may be some problems with the file names.
