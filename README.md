# Excel_interpolator
A desktop application to interpolate tables in excel

The file where the table is located can be inported by droping the file in. If you would like to type
the file name in simply uncheck the box labeled "I would like to Drop in file"

Choose the sheet that the table is on by typing the sheet name into the box labeled "Input Sheet Name"
future versions should detect sheet names from the incoming .xlsx file.

The application will automaticly look for numerical data and interpolate using the first column as the 
"x2" value in the interpolation formula "y2 = ((y3 - y1)/(x3-x1)) * (x2 - x1) + y1". y2 will be evaluated
at points inbetween x1 and x3 in steps of "Step Size". If a different using a column as "x" values is desired 
uncheck the box "interpolate from first column with data" and input the excel column to use in the "Column".
input example: "A" 

If writing interpolated table to a different sheet or file, simply uncheck the respective boxes and type
in the file path/sheet name. If no "Output Filepath" and/or "Output Sheet Name" are/is not found then
one will be generated.
