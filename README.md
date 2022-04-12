# Excel__Menu
加载定制Excel菜单_SmartView取数上传功能菜单

Public myName As name
Public MySheet As Worksheet
Public YesNo As Long

'Const Password = "Wildebeest!!"

Sub F_Open()
   
    On Error Resume Next
    application.CommandBars("AmyMenu").Delete
    application.CommandBars.Add(name:="AmyMenu", Position:=msoBarTop).Visible = True
                 
                 Set NEWBUTTON = CommandBars("AmyMenu").Controls.Add(Type:=msoControlButton, Before:=1)
                 With NEWBUTTON
                     .FaceId = 15
                     .Visible = True
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
                     .Caption = "EssUpload"
                     .Width = AutoFit
                     .OnAction = "UploadSelection"
'                     .OnAction = "RetrieveESS"
                 End With
                 
                 Set NEWBUTTON = CommandBars("AmyMenu").Controls.Add(Type:=msoControlButton, Before:=2)
                 With NEWBUTTON
                     .FaceId = 16
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
'                     .Caption = "DataExport"
                     .Caption = "MutiRetrive"
                     .Width = AutoFit
                     .OnAction = "MultSheet_SAPRetrieve"
'                     .OnAction = "DataExport"
                 End With
                  
                 Set NEWBUTTON = CommandBars("AmyMenu").Controls.Add(Type:=msoControlButton, Before:=3)
                 With NEWBUTTON
                     .FaceId = 17
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
                     .Caption = "ZoomIn"
                     .Width = AutoFit
                     .OnAction = "ZISelection"
'                     .OnAction = "DataImport"
                 End With
                 
                 Set NEWBUTTON = CommandBars("AmyMenu").Controls.Add(Type:=msoControlButton, Before:=4)
                 With NEWBUTTON
                     .FaceId = 18
                     .Visible = True
                     .Style = msoButtonIconAndCaption
                     .BeginGroup = True
                     .Caption = "EssSelection"
                     .Width = AutoFit
                     .OnAction = "RetrieveSelection"
                 End With
                 
End Sub
