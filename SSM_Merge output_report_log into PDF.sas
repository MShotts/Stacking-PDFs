
/*The following program performs the following steps:														*/
/*	1. Converts a GENASYS output .prn file into .pdf														*/
/*	2. Creates a report based on the contents of the GENASYS .prn file and saves it as .pdf					*/
/*		a.	For this example, the report is a listing of any run time messages in the output				*/
/*		b.	Please note that you can replace this report with anything of your choosing						*/
/*	3. Saves the log as .txt, reads it back into SAS, and then resaves as a .pdf							*/
/*	4. Builds and executes a Visual Basic script that calls in Acrobat Exchange to stack the 3 pdf as one	*/
/*	5. Deletes the temporary files used to build the stacked pdf											*/

/*Program was written by Matthew Shotts and presented on 2/4/2016 at the Skills Sharing Meeting				*/


%let newinput=SSM GEN Output for Merge into PDF.prn;	/*prn file for GEN Request	*/


/*Once this program is customized for your needs then there should be no need to make any changes beneath this line*/
options noxwait mprint;
%let abvinput=%scan(&newinput.,1,'.');



/*STEP 1 START*/
/*Creates a pdf out of the GENASYS .prn file*/
x "%str(%"C:\Program Files\Adobe\Acrobat 9.0\Acrobat\Acrodist.exe%" 
		/n/q %"B:\LR\Users\Matthew\SSM\&newinput.%" )";


/*Copies the pdf of the GENASYS .prn file to a staging folder*/
x "%str(copy %"\\ets\dfs\fs_stat_toeic\production\LR\Users\Matthew\SSM\&abvinput..pdf%" 
				%"\\ets\dfs\fs_stat_toeic\production\LR\Users\Matthew\SSM\pdf_staging\PDF_&abvinput..pdf%" )";
/*STEP 1 END*/



/*STEP 2 START*/
/*Read in the .prn file and define the heading and end of each page*/
Data All;
infile "C:\temp\&newinput."
lrecl=200 pad missover;
input 	@1	 All		$char200.
		@1	 endpg		$char7.
		@2	 Asterik	$char1.
		@3   AccNum		$char8.
		@4	 Head		$char130.
		@4   Desc		$char13.
;
Obs=_n_;
if Asterik="*" and substr(AccNum,1,1)=" " and Desc ne "Standardized " then Heading=strip(compress(Head,'*)')); 
if endpg="endPage" then Heading="endPage";
drop Asterik Head Desc;
run;

/*Retain the heading of each page until an end page is encountered*/
/*This allows the user to easily specify what records (page contents) to pull based on page heading*/
data AllwHeaders;
set All;
length new_Heading $100.;
if _n_=1 then new_Heading="Table of Contents";
retain new_Heading;
if Heading="" then new_Heading=new_Heading;
else new_Heading=Heading;

if substr(All,1,10)="(No errors" then call symput('SASNotes',PEobs);

drop Heading;
run;

/*Extract records that fell on pages with a header of 'Run Time Messages'*/
data RunTimePrep;
set AllwHeaders (keep=All new_Heading Obs);
length Snd_Heading $100.;

if Obs=1 then do;
Snd_Heading="";
end;

retain Snd_Heading;

if new_Heading in ("Run Time Messages","endPage") then Snd_Heading=new_Heading;
else Snd_Heading=Snd_Heading;
run;

/*Clean up the records so that the text of the warning is easily viewed*/
data RunTime;
set RunTimePrep;
where Snd_Heading="Run Time Messages";

Value=strip(compress(scan(scan(All,1,'('),1,')'),'*'));
rename Snd_Heading=Desc;

if Value in ('','Run Time Messages') then delete;
drop All Obs new_Heading;
run;

data RunTime2;
set AllwHeaders (keep=All new_Heading Obs);
where index(new_Heading,"Run Time")>0;
Value=strip(substr(All,2));
if substr(Value,1,1) in ("*",")") then delete;
rename new_Heading=Desc;
drop All Obs;
run;

data RunTimeOut; 
set RunTime RunTime2; 
run;


/*There are a number of ways to determine the number of records in a dataset, 													 */
/*however, I've found the following code to be the most reliable so far.														 */
/*I'm not entirely certain who deserves credit for this code as I've seen it appear in many online SAS support discussion threads*/
%macro callRT(z);

/*Opens the data RunTimeOut dataset created in the proc freq above.  If the dataset cannot be opened then DISD will equal 0.*/
%LET DSID=%SYSFUNC(OPEN(RunTimeOut,IN));

/*Retrieves the number of observations in the RunTimeOut dataset*/
%LET NOBS=%SYSFUNC(ATTRN(&DSID,NOBS));

/*This step is used to close the RunTimeOut dataset if it was actually opened (greater than 0)*/
%IF &DSID > 0 %THEN %LET RC=%SYSFUNC(CLOSE(&DSID));

/*The RunTimeOut dataset only includes records from pages with a heading of 'Run Time Messages'.  */
/*If there are no run Time Messages then we still want to give the user an assurance that the check was performed*/
%if &NOBS.=0 %then %do;

data RunTimeOut;
/*set RunTimeOut;*/
Desc="No Messages";
run;

%end;

%mend callRT;
%callRT (z);


/*Used the ods pdf destination to save the SAS output as a PDF*/
ods pdf file="\\ets\dfs\fs_stat_toeic\production\LR\Users\Matthew\SSM\pdf_staging\PDF_GEN Run Time Messages.pdf" style=minimal;
ods noptitle;

proc print data=RunTimeOut noobs;
title1 "The following Run Time Messages were found";
title2 "Investigate any error messages shown and either resolve them";
title3 "or type an explanation in the Excel file, QC Results for &nform.";
run;

ods pdf close;
/*STEP 2 END*/



/*STEP 3 START*/
/*There were some unexpected interaction wherein the output window was being saved rather than the log*/
/*I remedied this by saving the output first which seemed to clear the way for the log to save properly*/
%let outputFile = B:\LR\SAS\log\log_pdf.txt;
filename outputf "&outputFile.";
dm output 'file logout replace' output;

/*Creates a pdf of the log file*/
%let logFile = B:\LR\SAS\log\log_pdf.txt;
filename logOut "&logFile.";
dm log 'file logout replace' log;

ods listing close;

data readLog;
	infile logout truncover lrecl= 200;
	input @1 LogText $char200.;
run;

ods pdf file="\\ets\dfs\fs_stat_toeic\production\LR\Users\Matthew\SSM\pdf_staging\PDF_log.pdf" style=minimal;
ods noptitle;

proc print data=readLog noobs;
title1;
run;

ods pdf close;
/*STEP 3 END*/



/*STEP 4 START*/
/*The vast majority of the following code was found in this SAS article*/
/*http://analytics.ncsu.edu/sesug/2011/BB15.Welch.pdf*/

%let inDir=B:\LR\Users\Matthew\SSM\pdf_staging;
%let outDir=B:\LR\Users\Matthew\SSM;
%let NewFile= ;

********************;
*GET FOLDER CONTENTS;
********************;
*PART 1;
DATA prep;
folder = strip("&inDir");
rc = filename('files',folder);
/*dopen assigns a numeric value for which SAS can identify the directory*/
did = dopen('files');
/*dnum returns the number of members in the directory*/
numfiles = dnum(did);
iter = 0;
/*do loop to numfiles (number of files) so that the following steps are performed for each member of the directory*/
do i = 1 to numfiles;
/*dread returns the name of the directory member*/
text = dread(did,i);
if index(upcase(text),".PDF") then do;
iter + 1;
output;
end;
call symput("FileNum",put(iter,8.));
end;
/*dclose to close out the directory*/
rc = dclose(did);
RUN;
%put NUMBER OF PDF FILES TO STACK: %cmpres(&FileNum);

****************;
*BUILD VB SCRIPT;
****************;
*PART 2;
DATA indat1;
length code $150;
set prep;
code = 'Dim Doc'||compress(put(_n_,8.));
order = 1;
output;
code = 'Set Doc'||compress(put(_n_,8.))||'= CreateObject("AcroExch.PDDoc")';
order = 2;
output;
code = 'file'||compress(put(_n_,8.))||
' = Doc'||compress(put(_n_,8.))||
'.Open("'||strip("&inDir.\")||strip(text)||'")';
order = 3;
output;
RUN;
PROC SORT data = indat1;
by order;
RUN;

*PART 3;
DATA indat2;
length code $150;
set prep end = eof;
code = 'Stack = Doc1.InsertPages(Doc1.GetNumPages - 1, Doc'||
compress(put((_n_ - 1) + 2,8.))||
', 0, Doc'||
compress(put((_n_ - 1) + 2,8.))||
'.GetNumPages, 0)';
if eof then code = 'SaveStack= Doc1.Save(1, "'||
/*Define the location in which the stacked pdf will be saved*/
strip("&outDir.\")||
/*Define the name of the stacked pdf*/
strip("&abvinput.")||'wRep_Log'||'.pdf"'||')';
RUN;
********************;
*END BUILD VB SCRIPT;
********************;

*OUTPUT VB SCRIPT;
filename temp "&inDir.\PDFStack.vbs";
DATA allcode;
set indat1 indat2;
file temp;
put code;
RUN;

*RUN VB SCRIPT;
x "%str(%"&inDir.\PDFStack.vbs%" )";
/*STEP 4 END*/



/*STEP 5 START*/
/*The following code deletes the pdf_staging files as well as the temporary pdf and pdf conversion log files*/
filename delete pipe "Del /q &inDir.";
data _null_;
infile delete;
run;

x "%str(del %"\\ets\dfs\fs_stat_toeic\production\LR\Users\Matthew\SSM\&abvinput..pdf%" )";
x "%str(del %"\\ets\dfs\fs_stat_toeic\production\LR\Users\Matthew\SSM\&abvinput..log%" )";
/*STEP 5 END*/
