@echo off
%@Try%
  REM Normal code goes here
  python3 -u "../Attendence_Preload"
  py -u "../Attendence_Preload.py"
%@EndTry%
:@Catch
  REM Exception handling code goes here
:@EndCatch
echo Press any key to exit . . .
pause>nul