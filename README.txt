
Step	File				Why first?              				                (Estimated) Size (with comments)
1	formulas.py			Pure calculations; no workbook dependency.	                         3,000 lines 100% complete
2	validation.py		Dropdowns, validations, constants. Used everywhere.       			    3,112 lines 100% complete
3	workbook.py			Creates workbook, sheets, styles, tables. Depends on 1 & 2.	           5,008 lines 100% complete
4	create_match.py		Uses workbook helpers to generate match sheets.				    2,269 lines 100% complete
5	scorecard.py		Generates batting, bowling, innings summaries.				           2,400 lines 100% complete
6	leaderboard.py		Tournament standings and NRR calculations.				    4,000 lines 74% complete
7	score_ball.py		Core scoring engine. Uses almost every other module.		     5,700 lines written, 2,000 more to go
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~            current progress          ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
8	generate_workbook.py	Thin entry point that wires everything together.		                       3,000 lines 0% complete

Total progress percent: 83.600642855% (including comments)

Note: the current progress separator is placed below score_ball.py file since the file is more than half way (74%) done.


Varun@LAPTOP-U4ECJMSN MINGW64 ~/Desktop/GitHub Projects/cricket excel (main)
$ cloc .
       9 text files.
       9 unique files.
       2 files ignored.

github.com/AlDanial/cloc v 2.08  T=0.39 s (23.3 files/s, 68934.2 lines/s)
-------------------------------------------------------------------------------
Language                     files          blank        comment           code
-------------------------------------------------------------------------------
Python                           8           5927           9299          11336
Text                             1              5              0             27
-------------------------------------------------------------------------------
SUM:                             9           5932           9299          11363
-------------------------------------------------------------------------------

Lines of code written ~ 11,400 (25,489 including comments)
Lines of Code to be written ~ 5,000 including comments 
Project Progress percent based on lines of code including comments: 83.600642855%
