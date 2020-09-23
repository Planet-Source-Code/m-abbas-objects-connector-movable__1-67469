VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''
Option Explicit
Private Type point
    sXPos As Single
    sYPos As Single
End Type
Private Const DATALINE_SEP = 100
''''''' Returns Left Center POINT structure of the passed object
Private Function getLeftCenter(ByVal objToFind As Label) As point
getLeftCenter.sXPos = objToFind.Left
getLeftCenter.sYPos = (objToFind.Top + (objToFind.Top + objToFind.Height)) / 2
End Function
''''''' Returns Top Center POINT structure of the passed object
Private Function getTopCenter(ByVal objToFind As Label) As point
getTopCenter.sXPos = (objToFind.Left + (objToFind.Left + objToFind.Width)) / 2
getTopCenter.sYPos = objToFind.Top
End Function
''''''' Returns Right Center POINT structure of the passed object
Private Function getRightCenter(ByVal objToFind As Label) As point
getRightCenter.sXPos = objToFind.Left + objToFind.Width
getRightCenter.sYPos = (objToFind.Top + (objToFind.Top + objToFind.Height)) / 2
End Function
''''''' Returns Bottom Center POINT structure of the passed object
Private Function getBottomCenter(ByVal objToFind As Label) As point
getBottomCenter.sXPos = (objToFind.Left + (objToFind.Left + objToFind.Width)) / 2
getBottomCenter.sYPos = (objToFind.Top + objToFind.Height)
End Function
'''''''' Draw line with respect to the Object and the cursor arrow point
Public Function drawConnectLine(ByVal objPic1 As Label, ByVal sXPos As Single, ByVal sYPos As Single, ByVal iLineIndex As Integer, ByVal frm As Form)
''''''' Current X,Y is on the right side of the Object
    If sXPos > getRightCenter(objPic1).sXPos Then
        frm.lneLeft(iLineIndex).X1 = getRightCenter(objPic1).sXPos
        frm.lneLeft(iLineIndex).Y1 = getRightCenter(objPic1).sYPos
        frm.lneLeft(iLineIndex).X2 = (sXPos + getRightCenter(objPic1).sXPos) / 2
        frm.lneLeft(iLineIndex).Y2 = getLeftCenter(objPic1).sYPos
        
        frm.lneMiddle(iLineIndex).X1 = frm.lneLeft(iLineIndex).X2
        frm.lneMiddle(iLineIndex).Y1 = getLeftCenter(objPic1).sYPos
        frm.lneMiddle(iLineIndex).X2 = (sXPos + getRightCenter(objPic1).sXPos) / 2
        frm.lneMiddle(iLineIndex).Y2 = sYPos
        
        frm.lneRight(iLineIndex).X1 = frm.lneMiddle(iLineIndex).X2
        frm.lneRight(iLineIndex).Y1 = sYPos
        frm.lneRight(iLineIndex).X2 = sXPos
        frm.lneRight(iLineIndex).Y2 = sYPos
        '''''''Draw Arrow
        frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("RIGHT").Picture
        frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Tag = "RIGHT"
''''' Current X,Y is in the Leftside of the Object
   ElseIf sXPos < getLeftCenter(objPic1).sXPos Then
        
        frm.lneLeft(iLineIndex).X1 = getLeftCenter(objPic1).sXPos
        frm.lneLeft(iLineIndex).Y1 = getLeftCenter(objPic1).sYPos
        frm.lneLeft(iLineIndex).X2 = (getLeftCenter(objPic1).sXPos + sXPos) / 2
        frm.lneLeft(iLineIndex).Y2 = getLeftCenter(objPic1).sYPos
        
        frm.lneMiddle(iLineIndex).X1 = frm.lneLeft(iLineIndex).X2
        frm.lneMiddle(iLineIndex).Y1 = getLeftCenter(objPic1).sYPos
        frm.lneMiddle(iLineIndex).X2 = (sXPos + getLeftCenter(objPic1).sXPos) / 2
        frm.lneMiddle(iLineIndex).Y2 = sYPos
        
        frm.lneRight(iLineIndex).X1 = frm.lneMiddle(iLineIndex).X2
        frm.lneRight(iLineIndex).Y1 = sYPos
        frm.lneRight(iLineIndex).X2 = sXPos
        frm.lneRight(iLineIndex).Y2 = sYPos
        '''''''Draw Left Arrow
        frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("LEFT").Picture
        frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Tag = "LEFT"
 ''''' Current X,Y is at the Bottom of the Object
    ElseIf sXPos < getRightCenter(objPic1).sXPos _
        And sXPos > getLeftCenter(objPic1).sXPos Then
        If sYPos >= getBottomCenter(objPic1).sYPos Then
            frm.lneLeft(iLineIndex).X1 = getBottomCenter(objPic1).sXPos
            frm.lneLeft(iLineIndex).Y1 = getBottomCenter(objPic1).sYPos
            frm.lneLeft(iLineIndex).X2 = getBottomCenter(objPic1).sXPos
            frm.lneLeft(iLineIndex).Y2 = (getBottomCenter(objPic1).sYPos + sYPos) / 2
            
            frm.lneMiddle(iLineIndex).X1 = getBottomCenter(objPic1).sXPos
            frm.lneMiddle(iLineIndex).Y1 = frm.lneLeft(iLineIndex).Y2
            frm.lneMiddle(iLineIndex).X2 = sXPos
            frm.lneMiddle(iLineIndex).Y2 = (getBottomCenter(objPic1).sYPos + sYPos) / 2
            
            frm.lneRight(iLineIndex).X1 = sXPos
            frm.lneRight(iLineIndex).Y1 = frm.lneMiddle(iLineIndex).Y2
            frm.lneRight(iLineIndex).X2 = sXPos
            frm.lneRight(iLineIndex).Y2 = sYPos
            '''''''' Draw BOTTOM atrrow
            frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("BOTTOM").Picture
            frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
            frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
            frm.imgArrow(iLineIndex).Tag = "BOTTOM"
''''' Current X,Y is at the Top of the Object
        ElseIf sYPos < getTopCenter(objPic1).sYPos Then
            frm.lneLeft(iLineIndex).X1 = getTopCenter(objPic1).sXPos
            frm.lneLeft(iLineIndex).Y1 = getTopCenter(objPic1).sYPos
            frm.lneLeft(iLineIndex).X2 = getTopCenter(objPic1).sXPos
            frm.lneLeft(iLineIndex).Y2 = (getTopCenter(objPic1).sYPos + sYPos) / 2
            
            frm.lneMiddle(iLineIndex).X1 = getTopCenter(objPic1).sXPos
            frm.lneMiddle(iLineIndex).Y1 = frm.lneLeft(iLineIndex).Y2
            frm.lneMiddle(iLineIndex).X2 = sXPos
            frm.lneMiddle(iLineIndex).Y2 = (getTopCenter(objPic1).sYPos + sYPos) / 2
            
            frm.lneRight(iLineIndex).X1 = sXPos
            frm.lneRight(iLineIndex).Y1 = frm.lneMiddle(iLineIndex).Y2
            frm.lneRight(iLineIndex).X2 = sXPos
            frm.lneRight(iLineIndex).Y2 = sYPos
            ''''''''''''''''Draw TOP arrow
            frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("TOP").Picture
            frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
            frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
            frm.imgArrow(iLineIndex).Tag = "TOP"
''''' Current X,Y is on the Object
        Else
            frm.lneLeft(iLineIndex).X1 = sXPos
            frm.lneLeft(iLineIndex).Y1 = sYPos
            frm.lneLeft(iLineIndex).X2 = sXPos
            frm.lneLeft(iLineIndex).Y2 = sYPos

            frm.lneMiddle(iLineIndex).X1 = frm.lneLeft(iLineIndex).X2
            frm.lneMiddle(iLineIndex).Y1 = sYPos
            frm.lneMiddle(iLineIndex).X2 = sXPos
            frm.lneMiddle(iLineIndex).Y2 = sYPos

            frm.lneRight(iLineIndex).X1 = frm.lneMiddle(iLineIndex).X2
            frm.lneRight(iLineIndex).Y1 = sYPos
            frm.lneRight(iLineIndex).X2 = sXPos
            frm.lneRight(iLineIndex).Y2 = sYPos
            '''''''''''
            frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("LEFT").Picture
            frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
            frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
            frm.imgArrow(iLineIndex).Tag = "LEFT"
        End If
    End If
End Function
''''''''''''''''''Selecting the Line
'''Horz Line if y2-y1=0 and Vertical Line if x2-x1=0
'''Calc Rect area and check whether X and Y lies on that Rect
Public Function selectLine(ByVal frm As Form, ByVal sXPos As Single, ByVal sYPos As Single, Optional ByVal sTolerance As Single = 25) As Line
Dim sMaxTol As Single
Dim sMinTol As Single
Dim sX As Single
Dim sY As Single
Dim ctrl As Control
Dim iIndex As Integer

For Each ctrl In frm
    If TypeOf ctrl Is Line Then
        '''' vertical straight Line
        If ctrl.X2 - ctrl.X1 = 0 Then
            '''' Set max and Min tolerance to for Rect Area
            sMaxTol = ctrl.X1 + sTolerance
            sMinTol = ctrl.X1 - sTolerance
            If sYPos <= ctrl.Y2 And sYPos >= ctrl.Y1 And sXPos <= sMaxTol And sXPos >= sMinTol Then
                Set selectLine = ctrl
'                selectLine = ctrl.Name
'                iIndex = ctrl.Index
                Exit Function
            ElseIf sYPos <= ctrl.Y1 And sYPos >= ctrl.Y2 And sXPos <= sMaxTol And sXPos >= sMinTol Then
'                 selectLine = ctrl.Name
'                 iIndex = ctrl.Index
                 Set selectLine = ctrl
                 Exit Function
            End If
        ''''  Horizontal staright Line
        ElseIf ctrl.Y2 - ctrl.Y1 = 0 Then
            sMaxTol = ctrl.Y1 + sTolerance
            sMinTol = ctrl.Y1 - sTolerance
            If sXPos <= ctrl.X2 And sXPos >= ctrl.X1 And sYPos <= sMaxTol And sYPos >= sMinTol Then
'                selectLine = ctrl.Name
'                iIndex = ctrl.Index
                Set selectLine = ctrl
                Exit Function
            ElseIf sXPos <= ctrl.X1 And sXPos >= ctrl.X2 And sYPos <= sMaxTol And sYPos >= sMinTol Then
'                 selectLine = ctrl.Name
'                 iIndex = ctrl.Index
                 Set selectLine = ctrl
                 Exit Function
            End If
        End If
    End If
Next
Set selectLine = Nothing
Exit Function
End Function
''''''''''''''
Public Function moveConnectLine(ByVal object1 As Label, ByVal object2 As Label, ByVal iLineIndex As Integer, ByVal frm As Form, ByVal sSource As Boolean)
''''''' Object2 is on the left side of the Object1
 If getLeftCenter(object1).sXPos > getRightCenter(object2).sXPos Then
     frm.lneLeft(iLineIndex).X1 = getRightCenter(object2).sXPos
     frm.lneLeft(iLineIndex).Y1 = getRightCenter(object2).sYPos
     frm.lneLeft(iLineIndex).X2 = (getLeftCenter(object1).sXPos + getRightCenter(object2).sXPos) / 2
     frm.lneLeft(iLineIndex).Y2 = getLeftCenter(object2).sYPos
     
     frm.lneMiddle(iLineIndex).X1 = (getLeftCenter(object1).sXPos + getRightCenter(object2).sXPos) / 2
     frm.lneMiddle(iLineIndex).Y1 = getLeftCenter(object2).sYPos
     frm.lneMiddle(iLineIndex).X2 = (getLeftCenter(object1).sXPos + getRightCenter(object2).sXPos) / 2
     frm.lneMiddle(iLineIndex).Y2 = getLeftCenter(object1).sYPos
     
     frm.lneRight(iLineIndex).X1 = (getLeftCenter(object1).sXPos + getRightCenter(object2).sXPos) / 2
     frm.lneRight(iLineIndex).Y1 = getLeftCenter(object1).sYPos
     frm.lneRight(iLineIndex).X2 = getLeftCenter(object1).sXPos
     frm.lneRight(iLineIndex).Y2 = getLeftCenter(object1).sYPos
     
    If sSource = True Then
        frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("LEFT").Picture
        frm.imgArrow(iLineIndex).Left = (frm.lneLeft(iLineIndex).X1 + frm.lneLeft(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Top = (frm.lneLeft(iLineIndex).Y1 + frm.lneLeft(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Tag = "LEFT"
    Else
        frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("RIGHT").Picture
        frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Tag = "RIGHT"
    End If
     '''''''''''
''''''' Object2 is on the right side of the Object1
   ElseIf getRightCenter(object1).sXPos < getLeftCenter(object2).sXPos Then
    
    frm.lneLeft(iLineIndex).X1 = getLeftCenter(object2).sXPos
    frm.lneLeft(iLineIndex).Y1 = getLeftCenter(object2).sYPos
    frm.lneLeft(iLineIndex).X2 = (getLeftCenter(object2).sXPos + getRightCenter(object1).sXPos) / 2
    frm.lneLeft(iLineIndex).Y2 = getLeftCenter(object2).sYPos
    
    frm.lneMiddle(iLineIndex).X1 = (getLeftCenter(object2).sXPos + getRightCenter(object1).sXPos) / 2
    frm.lneMiddle(iLineIndex).Y1 = getLeftCenter(object2).sYPos
    frm.lneMiddle(iLineIndex).X2 = (getRightCenter(object1).sXPos + getLeftCenter(object2).sXPos) / 2
    frm.lneMiddle(iLineIndex).Y2 = getRightCenter(object1).sYPos
    
    frm.lneRight(iLineIndex).X1 = (getRightCenter(object1).sXPos + getLeftCenter(object2).sXPos) / 2
    frm.lneRight(iLineIndex).Y1 = getRightCenter(object1).sYPos
    frm.lneRight(iLineIndex).X2 = getRightCenter(object1).sXPos
    frm.lneRight(iLineIndex).Y2 = getRightCenter(object1).sYPos
    '''''''''''''
    If sSource = True Then
        frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("RIGHT").Picture
        frm.imgArrow(iLineIndex).Left = (frm.lneLeft(iLineIndex).X1 + frm.lneLeft(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Top = (frm.lneLeft(iLineIndex).Y1 + frm.lneLeft(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Tag = "RIGHT"
    Else
        frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("LEFT").Picture
        frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
        frm.imgArrow(iLineIndex).Tag = "LEFT"
    End If
''''''' Object2 is on the top or bottom of the Object1
  ElseIf getLeftCenter(object1).sXPos < getRightCenter(object2).sXPos _
        And getRightCenter(object1).sXPos > getLeftCenter(object2).sXPos Then
        If getTopCenter(object1).sYPos >= getBottomCenter(object2).sYPos Then
            frm.lneLeft(iLineIndex).X1 = getBottomCenter(object2).sXPos
            frm.lneLeft(iLineIndex).Y1 = getBottomCenter(object2).sYPos
            frm.lneLeft(iLineIndex).X2 = getBottomCenter(object2).sXPos
            frm.lneLeft(iLineIndex).Y2 = (getBottomCenter(object2).sYPos + getTopCenter(object1).sYPos) / 2
            
            frm.lneMiddle(iLineIndex).X1 = getBottomCenter(object2).sXPos
            frm.lneMiddle(iLineIndex).Y1 = (getBottomCenter(object2).sYPos + getTopCenter(object1).sYPos) / 2
            frm.lneMiddle(iLineIndex).X2 = getTopCenter(object1).sXPos
            frm.lneMiddle(iLineIndex).Y2 = (getBottomCenter(object2).sYPos + getTopCenter(object1).sYPos) / 2
            
            frm.lneRight(iLineIndex).X1 = getTopCenter(object1).sXPos
            frm.lneRight(iLineIndex).Y1 = (getBottomCenter(object2).sYPos + getTopCenter(object1).sYPos) / 2
            frm.lneRight(iLineIndex).X2 = getTopCenter(object1).sXPos
            frm.lneRight(iLineIndex).Y2 = getTopCenter(object1).sYPos
            '''''''''top
            If sSource = True Then
            
                frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("TOP").Picture
                frm.imgArrow(iLineIndex).Left = (frm.lneLeft(iLineIndex).X1 + frm.lneLeft(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Top = (frm.lneLeft(iLineIndex).Y1 + frm.lneLeft(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Tag = "TOP"
            Else
                frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("BOTTOM").Picture
                frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Tag = "BOTTOM"
            
            End If
        ElseIf getBottomCenter(object1).sYPos < getTopCenter(object2).sYPos Then
            frm.lneLeft(iLineIndex).X1 = getTopCenter(object2).sXPos
            frm.lneLeft(iLineIndex).Y1 = getTopCenter(object2).sYPos
            frm.lneLeft(iLineIndex).X2 = getTopCenter(object2).sXPos
            frm.lneLeft(iLineIndex).Y2 = (getTopCenter(object2).sYPos + getBottomCenter(object1).sYPos) / 2
            
            frm.lneMiddle(iLineIndex).X1 = getTopCenter(object2).sXPos
            frm.lneMiddle(iLineIndex).Y1 = (getTopCenter(object2).sYPos + getBottomCenter(object1).sYPos) / 2
            frm.lneMiddle(iLineIndex).X2 = getBottomCenter(object1).sXPos
            frm.lneMiddle(iLineIndex).Y2 = (getTopCenter(object2).sYPos + getBottomCenter(object1).sYPos) / 2
            
            frm.lneRight(iLineIndex).X1 = getBottomCenter(object1).sXPos
            frm.lneRight(iLineIndex).Y1 = (getTopCenter(object2).sYPos + getBottomCenter(object1).sYPos) / 2
            frm.lneRight(iLineIndex).X2 = getBottomCenter(object1).sXPos
            frm.lneRight(iLineIndex).Y2 = getBottomCenter(object1).sYPos
            
            If sSource = True Then
                frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("BOTTOM").Picture
                frm.imgArrow(iLineIndex).Left = (frm.lneLeft(iLineIndex).X1 + frm.lneLeft(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Top = (frm.lneLeft(iLineIndex).Y1 + frm.lneLeft(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Tag = "BOTTOM"
            Else
                frm.imgArrow(iLineIndex).Picture = frm.imlArrow.ListImages("TOP").Picture
                frm.imgArrow(iLineIndex).Left = (frm.lneRight(iLineIndex).X1 + frm.lneRight(iLineIndex).X2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Top = (frm.lneRight(iLineIndex).Y1 + frm.lneRight(iLineIndex).Y2) / 2 - (frm.imgArrow(iLineIndex).Width / 2)
                frm.imgArrow(iLineIndex).Tag = "TOP"
            End If
        End If
End If
End Function

