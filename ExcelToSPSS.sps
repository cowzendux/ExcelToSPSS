* Convert Excel files to SPSS
* By Jamie DeCoster

* This program requires a directory as the first argument, and will
* optionally take a second directory as the second argument.
* The program will convert both .xls and .xlsx files.

* Usage: ExcelToSPSS(source, [destination])
* The source is the location of the original Excel files, and
* the destination is where you want the SPSS files to be 
* placed. If the second directory is excluded, the program will
* put the SPSS files in a subdirectory off of the location of the
* original Excel files.  

**********
* Version History
**********
* 2013-06-08 Created
* 2014-03-19 Does not try to read temporary Excel files
* 2014-06-10 Increased delay between loading/saving each file

set printback = off.
begin program python.
import spss, os, time

def ExcelToSPSS(indir, outdir = "NONE"):

# Strip / at the end if it is present
    for dir in [indir, outdir]:
        if (dir[len(dir)-1] == "/"):
            dir = dir[:len(dir)-1]
            
# If outdir is excluded, create output directory if it doesn't exist
    if outdir == "NONE":
        if not os.path.exists(indir + "/SPSS"):
            os.mkdir(indir + "/SPSS")
        outdir = indir + "/SPSS"

# Get a list of all .xls files in the directory (xlsfiles)
    allfiles=[os.path.normcase(f)
    	for f in os.listdir(indir)]
    xlsfiles=[]
    for f in allfiles:
    	fname, fext = os.path.splitext(f)
    	if ('.xls' == fext and fname[:2] != '~$'):
    		xlsfiles.append(fname)

# Get a list of all .xlsx files in the directory (xlsxfiles)
    allfiles=[os.path.normcase(f)
    	for f in os.listdir(indir)]
    xlsxfiles=[]
    for f in allfiles:
    	fname, fext = os.path.splitext(f)
    	if ('.xlsx' == fext and fname[:2] != '~$'):
    		xlsxfiles.append(fname)

# Convert xls files
    for fname in xlsfiles:
        time.sleep(.50)
        submitstring = """GET DATA
  /TYPE=XLS
  /FILE='%s/%s.xls'
  /CELLRANGE=full
  /READNAMES=on
  /ASSUMEDSTRWIDTH=32767.
EXECUTE.
DATASET NAME $DataSet WINDOW=FRONT.""" %(indir,fname)
        spss.Submit(submitstring)

        time.sleep(.50)
        submitstring = """SAVE OUTFILE='%s/%s.sav'
  /COMPRESSED.
""" %(outdir,fname)
        spss.Submit(submitstring)

# Convert xlsx files
    for fname in xlsxfiles:
        time.sleep(.50)
        submitstring = """GET DATA
  /TYPE=XLSX
  /FILE='%s/%s.xlsx'
  /CELLRANGE=full
  /READNAMES=on
  /ASSUMEDSTRWIDTH=32767.
EXECUTE.
DATASET NAME $DataSet WINDOW=FRONT.""" %(indir,fname)
        spss.Submit(submitstring)

        time.sleep(.50)
        submitstring = """SAVE OUTFILE='%s/%s.sav'
  /COMPRESSED.
""" %(outdir,fname)
        spss.Submit(submitstring)

end program python.
set printback = on.
