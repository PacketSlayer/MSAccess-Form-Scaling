'Microsoft Access Class Module
Option Compare Database 'integer to long 12-31
Option Explicit

'//---BEGIN Public Class Properties
Public X As Integer 'x pos
Public Y As Integer 'y pos
Public h As Integer 'height
Public w As Integer 'width
Public fn As Integer 'font size
Public N As String 'Name of control
Public d As String 'datasheet height if applicable
'//---END Public Class Properties

'it may look like this is blank and void of code, but I assure you it is not.
'This is a container for a control on a form to keep and remember the original
'size and shape during resizing of form window. This will be added to an iterable
'collection thus remembering all controls on form
