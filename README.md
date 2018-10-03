# tsheets
Stata program that creates formatted tracking sheets

Stata help file: 
      
      Title
      
          tsheets -- Create formatted tracking sheets in Excel for field staff.
      
      
      Syntax
      
              tsheets varlist [if] [in], sortvars(varlist) [options]
      
          options               Description
          -------------------------------------------------------------------------
          Main
          * sortvars(varlist)   variables to sort tracking sheets
      
          Specifications
            titlevars(varlist)  variables for title of tracking sheets; if not
                                  specified, sortvars are used
            filename            specifies the name of the tracking sheets file;
                                  default is Tracking_Sheet
            number              number of entries on one page; default is 25
            variable            specifies that variable names should be used as
                                  column headers instead of variable labels
            nolabel             export variable values instead of value labels
          -------------------------------------------------------------------------
          * sortvars() is required.
      
      
      Description
      
          tsheets takes the dataset currently in memory and exports an excel file
          of specified tracking variables sorted by up to two variables.
      
          If only one sort variable is specified, tsheets will create a separate
          workbook for each of the values of the variable, with each sheet holding
          25 (or fewer) observations. If two sort variables are specified, each
          sheet of a workbook will be specific to the second sort variable.
      
      
      Remarks
      
          tsheets is meant to reduce the time spent creating tracking sheets for
          field officers by automatically formatting them to expand all cells and
          including titles, page numbers, and variable labels if specified.
      
      
      Examples
      
          Create tracking sheets for only the first three communities in comm_live
              . use master_data.dta
              . tsheets dob gender full_name phone1 phone2 contact1 if comm_live <=
                  3, sortvars(comm_live)
      
          Create a tracking sheet with an instruction for field officer to check
          box if they interview someone, with a name variable for the community
          comm_name
              . use master_data.dta
              . gen surveyed = ""
              . lab var surveyed "Check box if interviewed"
              . tsheets dob gender full_name phone1 phone2 contact1,
                  sortvars(comm_live) titlevars(comm_name)
      
          Tracking sheets for communities that are further stratified by villages,
          and both variables have name variables comm_name and vill_name
              . use master_data.dta
              . tsheets dob gender full_name phone1 phone2 contact1,
                  sortvars(comm_live vill_live) titlevars(comm_name vill_name) use
      
      
      Authors
      
          Rosemarie Sandino
          Christopher Boyer
      
          For questions or suggestions, submit a GitHub issue or e-mail
          researchsupport@poverty-action.org.
