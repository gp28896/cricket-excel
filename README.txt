
Step	File				Why first?              				                (Estimated) Size (with comments)
1	formulas.py			Pure calculations; no workbook dependency.	                         3,000 lines
2	validation.py		Dropdowns, validations, constants. Used everywhere.       			    3,112 lines
3	workbook.py			Creates workbook, sheets, styles, tables. Depends on 1 & 2.	           5,008 lines
4	create_match.py		Uses workbook helpers to generate match sheets.				    2,269 lines
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~            current progress          ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
5	scorecard.py		Generates batting, bowling, innings summaries.				           2,500 lines
6	leaderboard.py		Tournament standings and NRR calculations.				    2,500 lines
7	score_ball.py		Core scoring engine. Uses almost every other module.		                  3,500 lines
8	generate_workbook.py	Thin entry point that wires everything together.		                       1,000 lines


MINGW64 ~/Desktop/GitHub Projects/cricket excel (main)

$ cloc .
       7 text files.
       6 unique files.
       7 files ignored.

github.com/AlDanial/cloc v 2.08  T=0.20 s (30.2 files/s, 72419.8 lines/s)
-------------------------------------------------------------------------------
Language                     files          blank        comment           code
-------------------------------------------------------------------------------
Python                           5           3077           5910           5386
Text                             1              4              0             27
-------------------------------------------------------------------------------
SUM:                             6           3081           5910           5413
-------------------------------------------------------------------------------

Lines of code written ~ 5,413 (13,389 including comments)
Lines of Code to be written ~ 9,500 including comments 
Project Progress percent based on lines of code including comments: 58.50 %
