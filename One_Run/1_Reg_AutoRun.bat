@echo off
%@Try%
  REM Normal code goes here
  python3 -u "../Offline_Registration.py"
  py -u "../Offline_Registration.py"
%@EndTry%
:@Catch
  REM Exception handling code goes here
:@EndCatch
echo Press any key to exit . . .
pause>nul