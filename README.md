# STDFreadingutility
Parse STDF to be text (xml) and table (xlsx), also Ca/Cp/Cpk calculating

A. Under Python 3.7 environment, the script "Mstdfreader3.py" is STDF (ver4) reader which can parse the binary format to be text(xml) file. This script is revised from https://github.com/jasonshih/STDF_Reader and 'cahyo primawidodo 2016'.

Usage: one laptop with Win10 and Python 3.7 are installed.
C:\Users\User> python Mstdfreader3.py
Input STDF here, its filename (not zipped) should be file.stdf:  your stdf ( e.g. XYZ5678.stdf ) 

Result: as the attached “XYZ5678.xml”

B. Under Python 3.7 environment, the script “Mstdfxmltoxlsx2.py” can grep Parametric Test Record (PTR) and Multiple-Result Parametric Record (MPR), also collect its head/site/hard_bin/soft_bin/x_coord/y_coord of all tested parts from STDF xml file, then store into one table (xlsx).

Usage: one laptop with Win10 and Python 3.7 are installed.
C:\Users\User> python Mstdfxmltoxlsx2.py
Input xml (of STDF) here, its filename (not zipped) should be file.xml:  your xml (e.g. XYZ5678.xml) 

Result: as the attached “XYZ5678All.xlsx”

C. Under Python 3.7 environment, the script “Mstdfxmltoxlsx3.py” can calculate Ca/Cp/Cpk of PTR/MPR parametric results  from Good bins, then store into one table (xlsx).

Usage: similar to the above item B.

Result: as the attached “XYZ5678Good.xlsx”
