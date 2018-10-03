# tsheets
Stata program that creates formatted tracking sheets for field staff during field data collection.

Stata help file: 
      
<pre>
<b><u>Title</u></b>
<p>
    <b>tsheets</b> -- Create formatted tracking sheets in Excel for field staff.
<p>
<p>
<a name="syntax"></a><b><u>Syntax</u></b>
<p>
        <b>tsheets</b> <i>varlist</i> [if] [in]<b>,</b> <b><u>sort</u></b><b>vars(</b><i>varlist</i><b>)</b> [<i>options</i>]
<p>
    <i>options</i>               Description
    -------------------------------------------------------------------------
    Main
    * <b><u>sort</u></b><b>vars(</b><i>varlist</i><b>)</b>   variables to sort tracking sheets
<p>
    Specifications
      <b><u>title</u></b><b>vars(</b><i>varlist</i><b>)</b>  variables for title of tracking sheets; if not
                            specified, sortvars are used
      <b><u>file</u></b><b>name</b>            specifies the name of the tracking sheets file;
                            default is <i>Tracking_Sheet</i>
      <b><u>n</u></b><b>umber</b>              number of entries on one page; default is 25
      <b><u>var</u></b><b>iable</b>            specifies that variable names should be used as
                            column headers instead of variable labels
      <b><u>nolab</u></b><b>el</b>             export variable values instead of value labels
    -------------------------------------------------------------------------
    * <b>sortvars()</b> is required.
<p>
<p>
<b><u>Description</u></b>
<p>
    <b>tsheets</b> takes the dataset currently in memory and exports an excel file
    of specified tracking variables sorted by up to two variables.
<p>
    If only one sort variable is specified, <b>tsheets</b> will create a separate
    workbook for each of the values of the variable, with each sheet holding
    25 (or fewer) observations. If two sort variables are specified, each
    sheet of a workbook will be specific to the second sort variable.
<p>
<p>
<a name="remarks"></a><b><u>Remarks</u></b>
<p>
    <b>tsheets</b> is meant to reduce the time spent creating tracking sheets for
    field officers by automatically formatting them to expand all cells and
    including titles, page numbers, and variable labels if specified.
<p>
<p>
<a name="examples"></a><b><u>Examples</u></b>
<p>
    Create tracking sheets for only the first three communities in comm_live
        <b>. use master_data.dta</b>
        <b>. tsheets dob gender full_name phone1 phone2 contact1 if comm_live &lt;=</b>
            <b>3, sortvars(comm_live)</b>
<p>
    Create a tracking sheet with an instruction for field officer to check
    box if they interview someone, with a name variable for the community
    comm_name
        <b>. use master_data.dta</b>
        <b>. gen surveyed = ""</b>
        <b>. lab var surveyed "Check box if interviewed"</b>
        <b>. tsheets dob gender full_name phone1 phone2 contact1,</b>
            <b>sortvars(comm_live) titlevars(comm_name)</b>
<p>
    Tracking sheets for communities that are further stratified by villages,
    and both variables have name variables comm_name and vill_name
        <b>. use master_data.dta</b>
        <b>. tsheets dob gender full_name phone1 phone2 contact1,</b>
            <b>sortvars(comm_live vill_live) titlevars(comm_name vill_name) use</b>
<p>
<p>
<a name="authors"></a><b><u>Authors</u></b>
<p>
    Rosemarie Sandino
    Christopher Boyer
<p>
    For questions or suggestions, submit a GitHub issue or e-mail
    researchsupport@poverty-action.org.
</pre>
