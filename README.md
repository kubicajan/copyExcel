###INTRODUCTION###

This app was created to ease the work of copying from many excel sheets to another. With correct configuration, it can be a very powerful ally! Just put excel files into the projects directory and run it!

####Configuration###

The app itself does not need any configuration as it is right now. Should the need arise, here is how.
Most of the action happens in directory `services`. In both `ReadFromExcelImpl` and `WriteToExcelImpl` you will find a number of constants - these are used for configuring the app.

The most important ones from the `ReadFromExcelImpl` file are:
- *READ_FROM_FILE* - defines the name of the file to read from
- *MEASURE* - defines measure, which will be used for filtering values

-  *START_FROM_CELL_COORDINATES* - defines where the reading starts
-    *STOP_AT_CELL_COORDINATE* - defines where the reading ends

In case other batch of cells needs to be read:
-   *ADDITIONAL_START_CELL_COORDINATES* 
-   *ADDITIONAL_STOP_CELL_COORDINATES*

The most important ones from the `ReadFromExcelImpl` file are:

-    *WRITE_TO_FILE* - defines the name of the file to read from

App writes into two sheets:
-   *MMR_2_SHEET_NAME* - defines the MMR_2 sheet name
-   *FY21_SHEET_NAME* - defines the FY21 sheet name

-    *START_WRITING_FROM_CELL_COORDINATES_MMR_2* - defines where the writing starts for MMR_2 sheet
-    *START_WRITING_FROM_CELL_COORDINATES_FY21* - defines where the writing starts for FY21 sheet

-    *MEASURE_CELL_POSITION_IN_ROW* - measure cell position, for filtering purposes
-    *OPCO_CELL_POSITION_IN_ROW* - opco cell position, also for filtering purposes
    
    
####Conclusion####

If all of these constants are configured correctly, you no longer need to worry yourself with copying ever again! Simply run the program and it will do it for you.

####JAR####

In case you are tired of opening your IDEA, you can simply run the command `mvn clean package` and take your JAR executable file (from the `target` directory) with you! Simply put it in the same directory as your two excels, double click and the magic will happen.

Good luck!
