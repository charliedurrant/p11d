TCS Status Bar
==============

Custom Properties
PanelCount      r/w             Number of panels on status bar

Custom Methods
SetPanelWidths(%1,%2,...)       Set the widths of the panels
SetToolTips("Tip1","Tip2"...)   Set the tool tips of the panels
SetPB(PBmax)                    Set the maximum of the progress bar
SetStatus(progress, "Caption1","Caption2")      Set the progress value and the captions of the panels

The right most panel on the bar is the progress bar.  It will appear when you
call SetPB and disapear when SetStatus is called with progress > PBMax
