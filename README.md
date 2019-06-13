# MBTA-Slash-Line-Projects

## Purpose

The purpose of this project is to create a weekly report for Ray that gives "slashlines" by Buyer, Business Unite, purchase thresholds, and a specific Approver as well as two specific Buyers. Slashlines are compiled by looking at Buyer Backlogs from the most recent Monday as well as 4 weeks prior to that, and the slashline format is In/Out/Overlap/PlusMinus.

* "In" means Reqs that are in the more recent Backlog, but not the 4-week-old one
* "Out" means Reqs that are in the older Backlog, but not the more recent one
* "Overlap" means Reqs that are in both the older and newer Backlog
* "PlusMinus" is the count of "In" minus the count of "Out"

This project consists of 2 R scripts: one that produces most of the slashlines (`Slashline Report 040319.R`) and another that produces the slashlines for the single Approver (`PP Report 050119.R`).

### Package Requirements

#### Software Packages
* `R 3.5.1`

#### R Packages

* `openxlsx` for reading and writing excel files 
* `dplyr` for data manipulation
* `tidyverse` for data tidying

### Data Requirements

The first, main project takes 2 excel files as input, each the same but 28 days apart from each other. They are pulled from silverback's BACKLOG folder and emailed to onesself, with the file name format `BACKLOG-mmddyyyy-hhmmss.xlsx`. For ease of use, the files are renamed with the format below 

#### `BACKLOG-mmddyyyy.xlsx`

These excel files should contain the following columns, which should be edited in Excel to fit what is below (namely snake case):

| Columns                      | Purpose/Use                               |
| ---------------------------- | ----------------------------------------- |
| Business_Unit                | Business Unit assigned to that req        |
| Req_ID                       | Unique Identifier for each requisition    |
| Hold_From_Further_Processing | Hold Status of each req; either "Y" or "N"|
| Req_Date                     | Date of requisition's creation            |
| Origin                       | Not used                                  |
| Requester                    | Not used                                  |
| Approval_Date                | Date of requisitio's approval             |
| Buyer                        | Buyer assigned to that specific req       |
| Last_Change_Date             | Date that the req was last handled        |
| Buyer_Assignment             | Not used                                  |
| Hold_Req_Process             | Not used                                  |
| Out-to-Bid                   | Shows if req has been requested to go OTB |
| Req_Total                    | Dollar amount of the req                  |

### Output file(s)

This project will write a large number of Excel files, all of which are written to "C:/Users/NCAGLIUSO/Desktop/Slashline/Backlogs". All of the slashlines should be placed onto a sheet, with an additional slashline of just the NINV Buyer Team created in excel with counts from the full Buyer slashline. Three other sheets compile the workbook (`Slashline Report ddmmyy.xlsx`): Total data (all "In," "Out," and "Overlap"), Buyer Data ("Overlap" of the two specific Buyers) and PPAPAGNO data (explained below)

## Second Script: Individual Approver

The slashlines for the Approver, PPAPAGNO, follow the same format with the only difference being that the requisitions in his backlogs are those that need to be approved, as his role is different than that of a Buyer. In addition, it should be noted that three slashlines are compiled over three different date ranges: 7, 14, and 28 days. 

#### Software Packages

* `R 3.5.1`

#### R Packages

* `openxlsx` for reading and writing excel files 
* `dplyr` for data manipulation
* `tidyverse` for data tidying

### Data Requirements

The data for the individual Approver, PPAPAGNO, is also pulled from silverback in the file name format `APPROVER-BACKLOG-mmddyyyy-hhmmss.xlsx`. At the beginning the most recent Monday's file as well as the files from 7, 14, and 28 days ago needed to be pulled; however, this project has been going on long enough that only the most recent backlog has to be pulled, as the older ones are already on file off of silverback. As is seen above, the files are renamed once they are pulled for ease of use

#### `APPROVER-BACKLOG-mmddyyyy.xlsx`

These excel files should contain the following columns, which should be edited in Excel to fit what is below (namely snake case):

| Columns                      | Purpose/Use                               |
| ---------------------------- | ----------------------------------------- |
| User                         | Approver, which will alwasy be PPAPAGNO   |
| Status                       | Not used                                  |
| OperID                       | Not used                                  |
| Date_Time                    | Not used                                  |
| Origin                       | Not used                                  |
| Unit                         | Business Unit assigned to each PO         |
| Approval_Date                | Date of requisitio's approval             |
| PO_No                        | Unique Identifier for each PO             |
| Origin                       | Not used                                  |
| Merch_Amt                    | Not used                                  |
| Line                         | Line number of that specific PO           |

The script should be run three times, each time changing the file of the older Backlog from 7 to 14 and then 28 days ago.

### Output file(s)

After the script is run three times, the three slashlines should be added to the slashline sheet in the main workbook. In addition, the other data compiled (raw overlap data from each date range) should be added as the final sheet in the work book. It should be noted that the script writes the data to the same file every time, so the script should not be run three times successively. After each run, the data should be placed in the excel workbook as they are overwritten.
