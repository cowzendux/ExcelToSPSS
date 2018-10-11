# ExcelToSPSS

SPSS Python Extension function to convert all of the Excel files in a directory to SPSS data sets. The program will convert both .xls and .xlsx files.

## Usage
**ExcelToSPSS(indir, outdir)**
* "indir" is a required string argument indicating the directory of the original Excel files.
* "outdir" is an optional string argument indicating the location that the function should place the converted SPSS datasets. If is excluded, the program will put the SPSS files in a subdirectory named "SPSS" off of the location of the original Excel files.

## Example
**ExcelToSPSS(indir = "C:/Users/jamie/ICT study/Data/Excel",
outdir = "C:/Users/jamie/ICT study/Data/SPSS")**
* In this case, the function would take all of the .xls and .xlsx file found in the "C:/Users/jamie/ICT study/Data/Excel" directory, convert them to SPSS .sav files, and then save the new files in the directory "C:/Users/jamie/ICT study/Data/SPSS".
