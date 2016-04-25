# SQLTableOrViewToJSON
Take any MS SQL Server table or view, export the data as a JSON file

Author: Ahliana Byrd, 20151201
Purpose: To take any table or view, export the data as a JSON file

EXAMPLE: EXEC SQLTableOrViewToJSON 'MyTable', 'C:\WhereIStowMyJSON\'

WARNING: This code will enable xp_cmdshell
WARNING: This code will create 2 tables and leave them so they can be checked later, easier for debugging
WARNING: This code will create a saved stored procedure, SQLTableOrViewToJSON

This creates a string to declare variables, declare a cursor, walk the cursor, and output the resulting strings to a table
The string is then executed, so that the entire procedure is essentially done in a virtual memory space
The string that will be executed is printed in the Messages window, so if something doesn't seem to be working, 
	you can scroll down, grab the generated script, and paste it into a new window so you can debug it

Metadata code by: marc_s on StackOverflow, http://stackoverflow.com/questions/2418527/sql-server-query-to-get-the-list-of-columns-in-a-table-along-with-data-types-no
