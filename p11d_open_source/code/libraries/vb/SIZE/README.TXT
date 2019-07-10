The resize class automatically resizes controls on forms when the parent
form itself resizes. The class copes with controls on containers ie on 
frames. If the tag is LOCKBR for a control on a frame it will lock to 
the bottom right of the frame and not the parent form. There is one 
control this does not work for - data control. It has no container property
thus any tags which initiate resizing relative to an object edge will not resize
correctly when on a frame etc, only to the actual form.............

Using the resize class

------------------------------------------------------------------
First the two class modules must be added to the project

clsFormResize - clsfrmat.cls
clsCtrlAttibute - ctrlDisp.cls

Then for the form that requires control resizing:

------------------------------------------------------------------
in the form declarations

private mclsResize as new clsFormresize

in the form load

If Not (mClsResize.lInitResize(Me, L_DES_HEIGHT, L_DES_WIDTH)) then
  err.raise ERR_USER etc		
End If

------------------------------------------------------------------
L_DES_HEIGHT, L_DES_WIDTH represent the design time width and height.


------------------------------------------------------------------

The default resize option is scale height,width,top,left this is implemented when no text
is placed in the cntrl.tag

Specific Resize Options
-----------------------

In the tag of a control place

FREE = The deafault, scale relative and move relative.
LOCK = Lock the control, no move, no resize
EQUALISE = Retains design time distances from all form sides
EQUALISER = Retains design time distances from the right hand side of the window
EQUALISEB = Retains design time distances from the bottom side of the window
FREEEQUALISEBR = Moves,resizes but equalises to Bottom-Right
FREEEQUALISEB = get the gist   
LOCKBR = not resize but locks to bottom right, useful for buttons   
LOCKR = etc
FONT = scale the font
MOVEONLY = move only no scale
FREELOCKR = scales and moves but keeps design time distance from right
FREELOCKL = scales and moves but keeps design time distance from left
GRID = Scales grid columns accordingly
MOVEONLY = Move only nno scaling
FREELOCKR = Free movement\free scale but locks design time distance to the right
FREELOCKL = As above but to the left
FREELOCKTBL = Locks to top\bottom\left but scales width
FREELOCKTBR = As above but to right
FREELOCKTB = Locks to top bottom but will move left\right
FREELOCKTR = As above but to top\right
FREELOCKBR = As above but to bottom\right
LOCKBEQUALISER = Locks to bottom but keeps deign time distance from right
FREELOCKT = Locks to top but scales and moves left\right
FREELOCKB = As above but to bottom
CENTRE = Scales and positions in middle from left\right but not top\bottom
SCALEONLY
      

eg .tag = MOVEONLY
   .tag =FREE,FONT,GRID  (Moves scales font and changes grid size)

   Any number of options can be added by comma delimeting the items in
   the tag. Each is stored as a bit field and stored in a long.

The default is FREE, if you require FREE and eg. GRID the must type FREE,GRID


Methods,Properties

------------------------------------------------------------------

IsAdded(c as control) has it been added to the resize collection
AddCtrl(c as control) adds control to resize collection, required
	if adding controls at run time into control array
RemoveCtrl(c as control) removes control from resize collection

------------------------------------------------------------------

CAD

see demo project  q:\vb5\resize\demo for some examples


******** to do *************************************************************

i may have caused some problems when replacing error numbers with constants
wait for testing


add demo items to project




