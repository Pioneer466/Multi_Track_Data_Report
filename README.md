# Project: Student Performance Analytics CLI Application

This program is a full command-line analytics tool that processes academic performance data from multiple Excel files.
It allows users to explore statistics, generate visualizations, detect performance issues, and export reports — all directly from the terminal.

The tool automatically loads and cleans student data, computes detailed analytics, and provides an interactive CLI menu with colorful navigation.
All output files (graphics, exports…) are saved automatically in the output folder.

How to Run the Program :

1. Go to the project folder: (> cd 'file_path')

2. Install required libraries
=> pip install pandas seaborn matplotlib openpyxl xlsxwriter

3. Run the program
=>python3 report_generator.py

The interactive CLI will launch automatically in your terminal.

Extra Functionality :
	- Colored CLI interface for improved readability
	- Automatic performance alerts for low-performing tracks / cohorts

Files in the Project :
	- report_generator.py => Main python script (the program)
	- /data/ => Folder containing Excel files (student_grades_YYYY-YYYY.xlsx)
	- /output/ => Generated visuals & reports (AFTERWARDS WHEN LAUNCHING THE PROGRAM)
