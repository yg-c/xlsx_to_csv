SET "file=21.05.2021_Citrus.xlsx"
SET "restaurant=Citrus"
Set "month=05"
Set /A year=2021
SET /A start=22
SET /A end=31

for /l %%i in (%start%, 1, %end%) do IF %%i LSS 10 (copy %file% 0%%i.%month%.%year%_%restaurant%.xlsx) ELSE (copy %file% %%i.%month%.%year%_%restaurant%.xlsx) 