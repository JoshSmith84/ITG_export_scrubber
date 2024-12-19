========================
TPG ITG Export Scrubber
========================

Description
-------------
This program performs a standard preparation of
documentation exports for off-boarding clients. It takes a zip file of an export,
processes the csv files within, and outputs needed data into a single
xlsx workbook with worksheets.


Requirements
----------------
Functional Requirements:

* No valid data can be lost or skipped
* Output should be easily readable and organized
* Blank columns should be deleted automatically
* Cells with HTML should be changed to plain text
* A list of columns to always delete will be maintained and subject to change in the future
* Rows marked as archived should be deleted
* Configurations not marked as active should be deleted

Non-functional Requirements:

* Allow a user to choose either a single export, or a directory of exports
* If a directory, iterate over every zip file, and process only valid exports
* All output sheets should be formatted as a table
* Column width should allow the workbook to be opened and read with minimal re-sizing needed by the user

Functionality Not Required
---------------------------

At this time, the program does not need to delete any passwords



