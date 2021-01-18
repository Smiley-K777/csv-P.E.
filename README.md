# csv-P.E.
Position Exporting

Once VBScript is opend it will ask for the name of the file you would like the positions to be exported.

A few things that needs to be in place.
   1. File to be accessed has to be a {.csv} file.
   2. File must be in the same directory as the VBScript for it to be accessed.
   3. File must have Data in Cells A-K for each Row.

A little bit of how VBScript will run.
   1. VBScript will open your selected .csv file and read it.
   2. VBScript will have a set starting point of the 1st position. 
   3. Then VBScript will will run a loop reading Row by Row and looking for a difference in Cells B, C, D or E.
      3a. If Starting point is not the same as the end point {current Row that was read} it will create a copy 
          of the original file and delete all Rows that is not of that position.
      3b. It will then repeat this process untill the end point reaches the end of the file.

