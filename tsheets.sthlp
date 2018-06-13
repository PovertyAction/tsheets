{smcl}
{* *! version 1.0.0 Rosemarie Sandino 09jun2018}{...}
{title:Title}

{phang}
{cmd:tsheets} {hline 2}
Create formatted tracking sheets in Excel for field staff.


{marker syntax}{...}
{title:Syntax}

{p 8 10 2}
{cmd:tsheets} {it:{help varlist}} {help if:[if]} {help in:[in]}{cmd:,}
{opth sort:vars(varlist)} 
[{it:options}]

{* Using -help readreplace- as a template.}{...}
{* 20 is the position of the last character in the first column + 3.}{...}
{synoptset 20 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:Main}
{* Using -help heckman- as a template.}{...}
{p2coldent:* {opth sort:vars(varlist)}}variables to sort tracking sheets{p_end}

{syntab:Specifications}
{synopt:{opth title:vars(varlist)}}variables for title of tracking sheets; 
if not specified, sortvars are used{p_end}
{synopt:{opt file:name}}specifies the name of the tracking sheets file; 
default is {it:Tracking_Sheet}{p_end}
{synopt:{opt n:umber}}number of entries on one page; default is 25{p_end}
{synopt:{opt var:iable}}specifies that variable names should be used as column
headers instead of variable labels{p_end}
{synopt:{opt nolab:el}}export variable values instead of value labels{p_end}
{synoptline}
{p2colreset}{...}
{* Using -help heckman- as a template.}{...}
{p 4 6 2}* {opt sortvars()} is required.


{title:Description}

{pstd}
{cmd:tsheets} takes the dataset currently in memory and exports an excel 
file of specified tracking variables sorted by up to two variables.

{pstd}
If only one sort variable is specified, {cmd:tsheets} will create a separate 
workbook for each of the values of the variable, with each sheet holding 25 (or 
fewer) observations. If two sort variables are specified, each sheet of a 
workbook will be specific to the second sort variable. 


{marker remarks}{...}
{title:Remarks}

{pstd}
{cmd:tsheets} is meant to reduce the time spent creating tracking sheets 
for field officers by automatically formatting them to expand all cells and 
including titles, page numbers, and variable labels if specified. 


{marker examples}{...}
{title:Examples}

{pstd}
Create tracking sheets for only the first three communities in comm_live
{p_end}{cmd}{...}
{phang2}. use master_data.dta{p_end}
{phang2}. tsheets dob gender full_name phone1 phone2 contact1 if comm_live <= 3,
sortvars(comm_live){p_end}
{txt}{...}

{pstd}
Create a tracking sheet with an instruction for field officer to check box if they
interview someone, with a name variable for the community comm_name
{p_end}{cmd}{...}
{phang2}. use master_data.dta{p_end}
{phang2}. gen surveyed = ""{p_end}
{phang2}. lab var surveyed "Check box if interviewed" {p_end}
{phang2}. tsheets dob gender full_name phone1 phone2 contact1,
sortvars(comm_live) titlevars(comm_name){p_end}
{txt}{...}

{pstd}
Tracking sheets for communities that are further stratified by villages, and both
variables have name variables comm_name and vill_name
{p_end}{cmd}{...}
{phang2}. use master_data.dta{p_end}
{phang2}. tsheets dob gender full_name phone1 phone2 contact1,
sortvars(comm_live vill_live) titlevars(comm_name vill_name) use{p_end}
{txt}{...}


{marker authors}{...}
{title:Authors}

{pstd}Rosemarie Sandino{p_end}
{pstd}Christopher Boyer{p_end}

{pstd}For questions or suggestions, submit a
{browse "https://github.com/PovertyAction/tsheets/issues":GitHub issue}
or e-mail researchsupport@poverty-action.org.{p_end}

