The ucBrowse UsercControl is a Self-subclassed control which binds the 
Sytems Browse For Folder dialog and sets the FolderTree's parent window
to the current control.

Great care has been taken to ensure that the control is stable, but like
with all controls that are subclassed in this manor, there are a few things
that need to be observed to prevent issues.  

Known Issues:
1) Do not stop the application in the IDE with the End key (Button with Square Icon)!!
	
	[Side effects]
	The result may be a Usercontrol which is stuck in between Usercontrol
	and BFF shutdown. The BFF will not release and the application could
	act unresponsive. This only happens in the IDE and has not been observed
	in the past year of this controls use.

2) Do not attempt to Set the Root of the ucBrowse mutiple times in succession 
   without first calling the CloseUp method.

	[Side effects]
	This can cause the control to hang as it is being subclassed
 	repeatedly without first destroying the old window.....if this
	happens the BFF window is not released and the control acts
	as through it is hung.....at which point you must kill the
	process from the ProgramManager ;-(
	You have been notified....proceed at your own risk if you want
	to try this out!!

3) The CheckBoxes property is currently DesignTime only
	
	[Side Effects]
	This is an unfortunate side effect of changing the window style
	via bits. The control can be changed from CheckBoxes = False to
	CheckBoxes = True, but once the CheckBoxes are set the window style
	can not be reverted without destroying the window and recreating an
	instance. Repeated attempts at a work around have been unsuccessful, 
	so if you know of a way feel free to contact me. Also, there is no
	current "Nodes" support, and as such there is no current way to know
	which state icon has been applied (Checked / Unchecked) to provide
	feedback to the developer

4) The BackColor and ForeColor are currently not adjustable via the controls property

	[Side Effects] 
	As with the CheckBoxes properties these are set via style bits. The code to
	modify the style is in the control, but is not wired up at the current time.
	Consider this functionality under development for the time being, until
	I can figure out how to get the SysImageList icons to paint with transparent
	backcolors.

If you have questions / comments e-mail me at pwterrito@insightbb.com

Cheers,

Paul R. Territo, Ph.D
(TerriTop)