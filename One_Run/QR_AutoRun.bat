@echo off
%@Try%
  REM Normal code goes here
  python3 -u "../QR_Auto.py"
  py -u "../QR_Auto.py"
%@EndTry%
:@Catch
  REM Exception handling code goes here
:@EndCatch
echo Press any key to exit . . .
pause>nul
