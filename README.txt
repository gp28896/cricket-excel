
Step	File				Why first?              				                (Estimated) Size (with comments)
1	formulas.py			Pure calculations; no workbook dependency.	                         3,000 lines
2	validation.py		Dropdowns, validations, constants. Used everywhere.       			    3,112 lines
3	workbook.py			Creates workbook, sheets, styles, tables. Depends on 1 & 2.	           5,008 lines
4	create_match.py		Uses workbook helpers to generate match sheets.				    2,269 lines
5	scorecard.py		Generates batting, bowling, innings summaries.				           2,400 lines
6	leaderboard.py		Tournament standings and NRR calculations.				    4,000 lines
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~            current progress          ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
7	score_ball.py		Core scoring engine. Uses almost every other module.		                  4,000 lines
8	generate_workbook.py	Thin entry point that wires everything together.		                       2,000 lines


MINGW64 ~/Desktop/GitHub Projects/cricket excel (main)
cloc .
       8 text files.
       8 unique files.
       3 files ignored.

github.com/AlDanial/cloc v 2.08  T=0.28 s (28.9 files/s, 75652.1 lines/s)
-------------------------------------------------------------------------------
Language                     files          blank        comment           code
-------------------------------------------------------------------------------
Python                           7           4568           7651           8706
Text                             1              5              0             27
-------------------------------------------------------------------------------
SUM:                             8           4573           7651           8733
-------------------------------------------------------------------------------

Lines of code written ~ 8,700 (19,789 including comments)
Lines of Code to be written ~ 6,000 including comments 
Project Progress percent based on lines of code including comments: 76.73%
