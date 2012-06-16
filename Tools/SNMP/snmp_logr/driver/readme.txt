23 May 2010

These scripts were created to serve as a stand alone data driven SNMP logging tool.
All of the standard lmg library method were moved into this script and modified
to use only the parts that are neccessary and to consolidate into a single script.

The intent is to allow anyone use by simply install Ruby and Net-SNMP.
A future goal will be to use the Ruby snmp library

If many oids and few iterations are needed, the horizontal logger is the best choice
If many iterations are needed, the vertical logger is the best choice.


snmp_logr
 - horizontal logging

snmp_logr_1
 - general clean-up of the original script

snmp_logr_2
 - implement vertical logging - refactor script and spreadsheet to accomodate

snmp_logr_3
 - vertical logging
 - add column AutoFit feature (AutoSize columns after each row of data is written)

snmp_logr_4
 - vertical logging
 - Collect all get data into an array and write all at one to the spreadsheet
 - there was no performance improvement over snmp_logr_3

snmp_logr_5
 - vertical logging
 - collect all oids into an array and read all of the snmp data into an array 'data'.
 - write all of the snmp data (array 'data') into a range of cells.
 - negligible performance inprovement of 2 seconds


Dependencies
1) Ruby
2) Net-SNMP
3) Excel


Tested with:
1) ruby 1.8.6 p111
2) net-snmp 5.4.1
3) excel 2007



















