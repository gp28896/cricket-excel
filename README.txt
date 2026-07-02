
Step	File				Why first?												                        Estimated Size
1	formulas.py			Pure calculations; no workbook dependency.					        500–800 lines
2	validation.py		Dropdowns, validations, constants. Used everywhere.			    300–600 lines
3	workbook.py			Creates workbook, sheets, styles, tables. Depends on 1 & 2.	2,000–3,000 lines
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~            current progress          ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
4	create_match.py		Uses workbook helpers to generate match sheets.				    600–1,000 lines
5	scorecard.py		Generates batting, bowling, innings summaries.				      700–1,200 lines
6	leaderboard.py		Tournament standings and NRR calculations.					      700–1,200 lines
7	score_ball.py		Core scoring engine. Uses almost every other module.		    1,500–2,500 lines
8	generate_workbook.py	Thin entry point that wires everything together.		  100–200 lines


MINGW64 ~/Desktop/GitHub Projects/cricket excel (main)
$ cloc .
       5 text files.
       5 unique files.
       7 files ignored.

github.com/AlDanial/cloc v 2.08  T=0.18 s (27.9 files/s, 67480.7 lines/s)
-------------------------------------------------------------------------------
Language                     files          blank        comment           code
-------------------------------------------------------------------------------
Python                           4           2591           5138           4375
Text                             1              1              0              9
-------------------------------------------------------------------------------
SUM:                             5           2592           5138           4384
-------------------------------------------------------------------------------
