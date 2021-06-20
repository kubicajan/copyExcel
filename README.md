###INTRODUCTION###

This app was created to ease the work of copying from many excel sheets to another. With correct configuration, it can be a very powerful ally! Just put excel files into the projects directory and run it!

####Configuration###

The app itself does not need any configuration as it is right now. Should the need arise though, here is how.

The start of the whole process happens in the class `ServiceManager`. There, methods with configurations are initiated and 
together with them the whole process.

The program is able to copy from a set of `coordinates` - pairs of cells that indicate the start and the end of the block to be copied (be careful in which order you write the pairs though, it will be copied in that order).

Some sheets need to have the data `transposed`. Specify the name of the sheet and the starting cell and the program will take care of it. Same for the regular writing, of course.
    
####Conclusion####
If all of these constants are configured correctly, you no longer need to worry yourself with copying ever again! Simply run the program and it will do it for you.

####JAR####

In case you are tired of opening your IDEA, you can simply run the command `mvn clean package` and take your JAR executable file (from the `target` directory) with you! Simply put it in the same directory as your other excels, double click and the magic will happen.

Good luck!
