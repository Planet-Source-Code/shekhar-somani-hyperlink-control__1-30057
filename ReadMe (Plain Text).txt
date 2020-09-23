				HyperLink Control
				=================

    HyperLink control facilitates you to add web-like looking HyperLinks to
your VB forms easily. It gives very nice "Flat" look, which adds a different
style to your application. For example, in About box, you can add a hyperlink
to your Website or email address, just as available on any other web page...
HyperLink control will be a better choice than standard CommandButton.

    A really useful facility provided by this control is that it manages to
open the target by itself, unlike label, it is not only a display item, but
it also provides OpenTarget method, which opens the target Program, Document,
Web URL, or Email address in its default associated program using
ShellExecute API call.


More Features:
^^^^^^^^^^^^^^
    > Extended mouse events like MouseHover and MouseLeave makes it more 
      useful. 

    > Looks almost like any Web hyperlink, gives appearance facilities like
      hover font colors, Underline on mouse hover, etc. 

    > Provides necessary hand mouse pointers for available and not available
      link. 

--------------------------------------------------------------------------------

Programmer's Notes
^^^^^^^^^^^^^^^^^^

    The author of this control is Shekhar Somani. I live in India, and my
email address is given below:

Shekhar_Extreme@yahoo.com   -or-
Shekhar_d_s@yahoo.com

If you find any bugs or problems, or you have any comments, suggestions,
fixes, or anything to say about my work, please do mail me, I will really
appriciate.

    One problem that I am working on is the ToolTip, the control is not giving
tooltip properly, this seems because my control is not using its own property
for this, it relies upon VB's extender which automatically provides tooltip
property to every usercontrol, but since I am using API functions like
SetCapture and ReleaseCapture to provide extended mouse events, VB's extender
seems failing. So this needs to be worked, the application of this control
makes it almost compulsory that it should have ToolTip support. If anyone can
throw some light on this, please do. 

This control is a Freeware OpenSource, use it anywhere you like. It is free
for non-commercial uses. But any change in source, or distribution in public
requires author's prior permission.
--------------------------------------------------------------------------------

			HyperLink Control Reference
			===========================

Properties
~~~~~~~~~~
	- Target 
	- LinkAvailable 
	- UnderlineOnHover 
	- HoverForeColor 
	- UseHoverForeColor 
	
Methods
~~~~~~~
	OpenTarget  

--------------------------------------------------------------------------------
"Target" Property (String)
~~~~~~~~~~~~~~~~~~~~~~~~~~

    Specifies the path of the target of the hyperlink, this can be any of
these:

	An application (*.exe *.com *.bat etc.) 
	A Document (*.*) - Uses the default associated application to open 
	Internet URL (Tries to open in default program, so it needs IE4 update) 
	Email Address (Tries to open in default program, so it needs IE4 update) 

Note: To open the target, OpenTarget method can be called anytime, but this
property must be set to something before calling the OpenTarget method,
otherwise the control throws an error.

--------------------------------------------------------------------------------
LinkAvailable (Boolean)
~~~~~~~~~~~~~~~~~~~~~~~

    This Boolean property acts as simple flag which can be set to True/False
by the programmer to indicate the control that the link is been disabled
temporarily.

    The noticeable change that this property makes is that once set to False,
the control shows the "Not Available" mouse icon and any call to OpenTarget
method raises an error. 

--------------------------------------------------------------------------------
UnderlineOnHover (Boolean)
~~~~~~~~~~~~~~~~~~~~~~~~~~

    If set to True, the control underlines the caption text as soon as the
mouse moves over the control, and resets it as the mouse leaves the control
area. This property makes it look like a real hyperlink. Default setting for
this property is True.

--------------------------------------------------------------------------------
HoverForeColor (Color)   &   UseHoverForeColor (Boolean)
~~~~~~~~~~~~~~~~~~~~~~       ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    If UseHoverForeColor is set to True, the control changes the caption
color to the color set in HoverForeColor property, which is reset when mouse
leaves the control area. 

Default Values:
	HoverForeColor = vbBlue 
	UseHoverForeColor = True  

--------------------------------------------------------------------------------
OpenTarget Method
~~~~~~~~~~~~~~~~~

Parameters: (All Optional)
	[OwnerHWnd]  Long 
	[WindowStyle]  APIWindowStyleConstants 
	[StartupDirectory]  String 

    Returns: Long value returned by ShellExecute API call

    This method opens the target of the hyperlink by calling the ShellExecute
API function and returns appropriate value. It has three optional parameters,
which provide better control over the application called.

    Note: This method fails and raises an error if LinkAvailable property is
set to False, or Target property is not set to anything.
--------------------------------------------------------------------------------

HyperLink Control Readme

Shekhar Somani
23rd December, 2001
