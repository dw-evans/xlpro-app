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

3. `HIGH` `WAITING FOR OTHER ASPECTS TO BE RESOLVED` Record videos of usage for website

6. `LOW` Implement relative Paths for path objects.
   1. All Path-typed variables can have the option to be relative to the workbook Path(\_\_file\_\_).parent.parent

7. `Website`
   1. `MED` Fix home page hamburger menu breaking
   2. `HIGH` Fix home page content
   3. `HIGH` Add home page assets (4x videos)
   4. `HIGH` Rework 'Getting Started' slash tutorial 

8. `BUG` `MED` Debug errors with modules failing over time

11. `COMPLETED` `HIGH` Develop way to push requirements
    1.  Start with a function to push reqs.txt. This should warn that the files (reqs.txt, .python-version may be different, and the user should compare.)
    2.  Git should probably be used inherently
    3.  `COMPLETED` for requirements txt
    4.  `COMPLETED` `BUG` for python version!

13. `HIGH` Test reloading a configured environment on multiple machines 

21. `PARTIAL` `HIGH` Investigate undo-ing
    1.  Test implementing a synchronous calculation mode which executes each function sequentially.
        1.  `INCOMPLETE` Async work done but not 
    2.  Investigate doing workbook backups. Could save the workbook e.g. ./.xlpro/Book1.xlsm.2025-06-04_1134.bak. Timeframe between backups could be set in the user's xlpro configuration file. Button to revert to previous backups. Could be a console menu with several selections
    3. `POSSIBLE OPTION` Write a custom action tracker based on user events
       1. Track all excel actions so we can at least track typing events
       2. Write a xlpro.xlam sub that can be called by xlpro as an atomic formula resetter without event tracking
    4. `COMPLETED` VBA-based undo tracker. Todo link up to some buttons.

22. `EXPERIMENTAL` `MED` Asynchronous functions
    1.  yielding functions can be decorated with a timer (breaks undo stack) 
    2.  hook onto existing rtd functionality by spinning up another com process
    3.  `BUG` Developed the functional code for this but only works after the first calculation.
23. `MED` test behaviour for external files. 
    1.  Can external files be registered safely? I would guess this would not play nicely with the sys module hack.
24. `DONE` `HIGH` Investigate trimming com calls by running an argument check before execution  

25. `OBSOLETE` `MED` Investigate configuring xlpro-cli as a server which can receive commands.
26. 
27.  `HIGH` Just investigate moving the heavy imports to the heavier commands and not the lightweight ones FOR XLPRO-CLI to improve performance

29.  `HIGH` Subroutines development
    
31.  `EXPERIMENTAL` Subroutine configuration via function?

32.  `PARTIAL` `MED` Help
     1.   Add arguments to workbook namespace so they can be looked up from the workbook
     2.   Implementation done

33.  `HIGH` Default arguments

34.  `HIGH` Re-investigate jsonified arguments, probably change to something that can handle n by 2 array key value pairs instead of a json string.

36.  `MED` Add `@vectorize` decorator similar to the jsonify decorator
     1.   @vectorize(((arg_name_to_vectorise1, new_type1), (arg_name_to_vectorise2, new_type2)))

39.   `FEATURE` Yield-timer functions

40.   

41.   

42.   

43.   

44.   

45.   