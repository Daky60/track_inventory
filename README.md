# Inventory sorter
An automated script that attempts to find logged on users on specified computers and then emails said users.  
In addition to this, it will spit out which users has been found on said computers to foundusers.csv  
Ideally you should also automatically feed this script with a list of computers you're looking for. (flagged.csv)  

## My use case
I put this together to a) familarize myself with POSH and b) track inventory with discrepancies.  
It is up to you what you feed the program.

## Warning
It relies on a 2 hour interval between runs or it will send multiple emails to logged in users.  
Change on line 76 if you want to run it more frequently. (not recommended)
