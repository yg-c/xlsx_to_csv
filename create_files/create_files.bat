SET "file=01.05.2021_Désalpe.xlsx"
SET /A start=12
SET /A end=31

for /l %%i in (%start%, 1, %end%) do IF %%i LSS 10 (copy %file% 0%%i.05.2021_Désalpe.xlsx) ELSE (copy %file% %%i.05.2021_Désalpe.xlsx) 