IF X%1X==XX goto error
IF X%2X==XX goto error
XDEL %WINSYS%\%1
XDEL %2\%1
goto finished
:error
@ECHO usage: %0 dllname release-code-path
:finished