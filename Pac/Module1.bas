Attribute VB_Name = "Module1"
Option Explicit

Public Function getAppColor(aKey As String) As Long

    ' This is a dummy function until we release the code for our custom
    '  "theme resource" file, which stores colors and bitmaps

    Select Case LCase(aKey)
        Case "body"
            getAppColor = RGB(58, 110, 165)
        Case "selected"
            If gbCustomTexture Then
                getAppColor = RGB(242, 162, 153)
            Else
                getAppColor = RGB(186, 186, 204)
            End If
        Case "selectedtext"
            getAppColor = vbBlack
        Case "generaltext"
            getAppColor = vbBlack
        Case "bordercolor"
            If gbCustomTexture Then
                getAppColor = RGB(240, 72, 72)
            Else
                getAppColor = RGB(85, 85, 118)
            End If
        Case "table1bg"
            getAppColor = RGB(223, 223, 223)
        Case "table2bg"
            getAppColor = RGB(241, 241, 241)
        Case "headingbg"
            getAppColor = RGB(128, 128, 166)
        Case "headingtext"
            getAppColor = RGB(231, 231, 255)
        Case "menubg"
            getAppColor = RGB(58, 110, 165)
        Case "MenuText"
            getAppColor = vbBlack
    End Select

End Function

