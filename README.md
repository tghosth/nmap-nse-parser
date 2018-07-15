# nmap-nse-parser

This is a very basic script to parse NSE results out of a `.nmap` file and present them in a tabular form alongside IP addresses. You can see the options by running the script without parameters.

When you open the file in Excel for the first time, for some reason, the new lines do not display correctly.

I write a quick and very dirty [VBA script](fixnewlines.vba) that fixes this issue if you select all the cells in the table that was written and run the VBA script. It just simulates an edit in each cell which triggers the newlines to display properly.
