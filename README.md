# xlpro Application

## Modules

### xlpro_module

- Python module which manages python server creation.
- Wheel file is created for internal distribution to the installer

### xlpro_cli

- Command line tool which manages workbook and virtual environment configuration
- Generates data within the installation directory.


### xlpro_installer

- Installer for distribution
- Includes the following:
  - uv.exe (indirectly fetched)
  - xlpro.xlam (xlpro-add in file, copy moved to xlstart)
  - xlpro*.whl (wheel file for building source for each virtual environment)
  - xlpro-cli.exe (cli tool added to path to work with xlpro.xlam)
  - envs/ (folder for storing independent virtual environments)
  - *configuration files and other*

### xlpro_addin

- Compiles the excel add-in file xlam

### xlpro_examples

- Example spreadsheeets and code

## Development Notes

1. `HIGH` Stability testing - In progress
   1. Seems pretty good as of now

2. `HIGH` Example development
   1. `HIGH` `Functions`
      1. `MED` `Matplotlib` plotting examples
      2. `MED` `Seaborn` plotting examples
      3. `MED` `Bokeh` plotting examples
      4. `HIGH` Data-focused example - dataframe handling, file loading
      5. `LOW` AI example
      6. `MED` Financial modelling

   2. `HIGH` `Subroutines`
      1. Excel COM interop subroutine (may run into stability issues due to multiple threads accessing trying to access Excel at the same time -> requires a scheduler, e.g. sending a callable to the client manager's queue. If its a callableWrapperClass execute the function...? Possible to fetch the result)
      2. Non-Excel subroutine
      
   3. `EXPERIMENTAL` Conditional formatting

3. `HIGH` `WAITING FOR OTHER ASPECTS TO BE RESOLVED`Record videos of usage for website

4. `DONE` `HIGH` Tidy calculation cycle - Eliminate COM calls for promised arguments etc within VBA.
   1. At least can mitigate the non-array data, checking all the nested values of an array is probably not worthwhile.
   2. Estimated cost overlap for an out of process COM call versus checking values would be around 6000 - could implement a threshold value

5. `LOW` Implement `@vectorize` decorator similar to the jsonify decorator.

6. `LOW` Implement relative Paths for path objects.
   1. All Path-typed variables can have the option to be relative to the workbook Path(\_\_file\_\_).parent.parent

7. `Website`
   1. `MED` Fix home page hamburger menu breaking
   2. `HIGH` Fix home page content
   3. `HIGH` Add home page assets (4x videos)
   4. `HIGH` Rework 'Getting Started' slash tutorial 

8. `BUG` `MED` Debug errors with modules failing over time

9.  `DONE` `BUG` `HIGH` Investigate how to prevent shell window freezing and needing an enter press

10. `PARTIAL` `HIGH` Sync up serve/start so they behave the same.
    1.  Serving can detect the existing process and just ignore the request
    2.  Serving is followed up by issuing a start call
    3.  Start call probably needs to be captured using checks on stdout pipe text content for trigger.
    4.  `INCOMPLETE` Implement close server button



11. `COMPLETED` `HIGH` Develop way to push requirements
    1.  Start with a function to push reqs.txt. This should warn that the files (reqs.txt, .python-version may be different, and the user should compare.)
    2.  Git should probably be used inherently
    3.  `COMPLETED` for requirements txt
    4.  `COMPLETED` `BUG` for python version!

12. `LOW` Git integration

13. `HIGH` Test reloading a configured environment on multiple machines 

14. `LOW` Add .xlpro-version file to ./Book1.xlsx.xlpro directory


15. `DONE` `MED` non-destructive launch.json writing
    1.  Prototyped, not tested - pushed
    
16. `DONE` `MED` non-descrutcute settings.json writing
    1.  Prototyped - looks ok

17. `PARTIAL` `HIGH` pushing requirements on workbook save
    1.  Prototyped function and button

18. `DONE` `HIGH` Editing config from excel button 
    1.  Prototyped, not tested.
    2.  Looks ok

19. `DONE` `MED` Investigate lockfile behaviour, integrate sync and start? Create Close button/Function?
    1.  
20. `PARTIAL` `MED` Close Server function in xlpro-cli
    1. Force restart on serve() button press. User still needs to manually close

21. `HIGH` Investigate undo-ing
    1.  Test implementing a synchronous calculation mode which executes each function sequentially.
        1.  `INCOMPLETE` Async work done but not 
    2.  Investigate doing workbook backups. Could save the workbook e.g. ./.xlpro/Book1.xlsm.2025-06-04_1134.bak. Timeframe between backups could be set in the user's xlpro configuration file. Button to revert to previous backups. Could be a console menu with several selections
    3.  `POSSIBLE OPTION` Write a custom action tracker based on user events
22. `EXPERIMENTAL` `MED` Asynchronous functions
    1.  yielding functions can be decorated with a timer (breaks undo stack) 
    2.  hook onto existing rtd functionality by spinning up another com process
    3.  `BUG` Developed the functional code for this but only works after the first calculation.
23. `MED` test behaviour for external files. 
    1.  Can external files be registered safely? I would guess this would not play nicely with the sys module hack.
24. `DONE` `HIGH` Investigate trimming com calls by running an argument check before execution  

25. `MED` Investigate configuring xlpro-cli as a server which can receive commands.
     1.   Just investigate moving the heavy imports to the heavier commands and not the lightweight ones

26.  `HIGH` configure xlpro-cli to ensure that VBIDEReference is loaded as a reference for the workbook!
27.  `HIGH` investigate vbe control options

28.  `HIGH` Subroutine
29.  `HIGH` Button development
30.  `EXPERIMENTAL` Subroutine configuration via function?

31.  `MED` Help
     1.   Add arguments to workbook namespace so they can be looked up from the workbook

32.  `HIGH` Default arguments
33.  `HIGH` Re-investigate jsonified arguments, probably change to something that can handle n by 2 array key value pairs instead of a json string.
34.  


35.  `HIGH` Improve performance
     1.   `LOW` Implement threadpoolexecutor for performance
     2.   `DONE` Pre process arguments in vba
     3.   Add `@vectorize` decorator similar to the jsonify decorator

36.  

37.  