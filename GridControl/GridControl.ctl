VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl GridControl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "GridControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_BackStyle = 1
Const m_def_RowCount = 0
Const m_def_SelectedCount = 0
'Property Variables:
Dim m_BackStyle As Integer
Dim m_RowCount As Long
Dim m_SelectedCount As Long
'Event Declarations:
Event Click() 'MappingInfo=listview1(1),listview1(1),-1,Click
Event DblClick() 'MappingInfo=listview1(1),listview1(1),-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=listview1(1),listview1(1),-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=listview1(1),listview1(1),-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=listview1(1),listview1(1),-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=listview1(1),listview1(1),-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=listview1(1),listview1(1),-1,MouseUp
Event ColumnClick(ByVal ColumnHeader As ColumnHeader) 'MappingInfo=listview1(1),listview1(1),-1,ColumnClick
Event ItemClick(ByVal Item As ListItem) 'MappingInfo=listview1(1),listview1(1),-1,ItemClick

Dim TooltipCol_1 As Long
Dim TooltipCol_2 As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''  Multi Line Tooltip for ListView Control  ''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    
    Const LVM_FIRST = &H1000&
    Const LVM_HITTEST = LVM_FIRST + 18
    
    Private Type POINTAPI
        x As Long
        y As Long
    End Type
    
    Private Type LVHITTESTINFO
       pt As POINTAPI
       flags As Long
       iItem As Long
       iSubItem As Long
    End Type
    
    Dim TT As CTooltip
    Dim m_lCurItemIndex As Long

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
On Error Resume Next
    BackColor = ListView1(1).BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    ListView1(1).BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
On Error Resume Next
    ForeColor = ListView1(1).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    ListView1(1).ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
On Error Resume Next
    Set Font = ListView1(1).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set ListView1(1).Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
On Error Resume Next
    BorderStyle = ListView1(1).BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    ListView1(1).BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a form or control."
    ListView1(1).Refresh
End Sub


Private Sub ListView1_Click(Index As Integer)
    RaiseEvent Click
End Sub

Private Sub ListView1_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Private Sub listview1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub listview1_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub listview1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub ListView1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub


Private Sub listview1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim lvhti As LVHITTESTINFO
    Dim lItemIndex As Long
    
    If TooltipCol_1 <> 0 Or TooltipCol_2 <> 0 Then
        'Show multi-line tooltip
        
        lvhti.pt.x = x / Screen.TwipsPerPixelX
        lvhti.pt.y = y / Screen.TwipsPerPixelY
        lItemIndex = SendMessage(ListView1(1).hWnd, LVM_HITTEST, 0, lvhti) + 1
        
        If m_lCurItemIndex <> lItemIndex Then
            m_lCurItemIndex = lItemIndex
            If m_lCurItemIndex = 0 Then   ' no item under the mouse pointer
                TT.Destroy
            Else
                If TooltipCol_1 = 0 Then
                    TT.Title = "" 'it doesn't need a title
                Else
                    TT.Title = GetGridData(m_lCurItemIndex, TooltipCol_1)
                End If
                
                If TooltipCol_2 = 0 Then
                    TT.TipText = "" 'it needs to have text
                Else
                    TT.TipText = GetGridData(m_lCurItemIndex, TooltipCol_2)
                End If
                TT.Create ListView1(1).hWnd
            End If
        End If
    
    End If
    
End Sub


Private Sub listview1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,Checkboxes
Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the list."
On Error Resume Next
    Checkboxes = ListView1(1).Checkboxes
End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)
    ListView1(1).Checkboxes() = New_Checkboxes
    PropertyChanged "Checkboxes"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=listview1(1),listview1(1),-1,MultiSelect
Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
On Error Resume Next
    MultiSelect = ListView1(1).MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    ListView1(1).MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub SortByColumn(pColumn As Long, pAscending As Boolean)
Attribute SortByColumn.VB_Description = "Sorts the grid by the specified column."
    
    If pAscending = True Then
        ListView1(1).SortOrder = lvwAscending
    Else
        ListView1(1).SortOrder = lvwDescending
    End If
    
    ListView1(1).SortKey = pColumn - 1
    ListView1(1).Sorted = True
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub SetToolTips(pColumn1 As Long, pColumn2 As Long)
Attribute SetToolTips.VB_Description = "Sets the tooltips for the grid."
     
    TooltipCol_1 = pColumn1
    TooltipCol_2 = pColumn2
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub SetColumnWidth(pColumnNum As Long, pColumnWidth As Long)
Attribute SetColumnWidth.VB_Description = "Sets the width of the specified column."
    
    ListView1(1).ColumnHeaders(pColumnNum).Width = pColumnWidth

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get RowCount() As Long
Attribute RowCount.VB_Description = "Returns the number of rows."
Attribute RowCount.VB_MemberFlags = "400"
    RowCount = m_RowCount
End Property

Public Property Let RowCount(ByVal New_RowCount As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_RowCount = New_RowCount
    PropertyChanged "RowCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0
Public Function LoadGrid(pSQL As String, pConnectionString As String, pNoDataMsg As String, pNoDataCol As Long) As Boolean
Attribute LoadGrid.VB_Description = "Sets the contents of the grid. Returns False if no records returned from SQL."

    'pSQL               the SQL string to pass to the database
    'pConnectionString  the connection string used to get to the database
    'pNoDataMsg         the message you want displayed in the 1st line of the grid if no
    '                   records are returned from the SQL statement
    'pNoDataCol         the column you want the above message to appear in
    '                   For Example: Usually 1, but could be 2 if you want to set Column 1's width to zero
    
    Dim curCol As Long
    Dim i As Long
    Dim maxRow As Long
    Dim maxCol As Long
    Dim rs As ADODB.Recordset
    
    'List views don't do well at refreshing themselves,
    'so this is the work-around
    Unload ListView1(1)
    Load ListView1(1)
    ListView1(1).Height = UserControl.Height
    ListView1(1).Width = UserControl.Width
    
    ListView1(1).Visible = True
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    'Execute the SQL
    rs.Open pSQL, pConnectionString
    
    'Add the column headings
    For curCol = 1 To rs.Fields.Count
        ListView1(1).ColumnHeaders.Add curCol, , rs.Fields(curCol - 1).Name
    Next curCol

    If rs.EOF Then
        'No records returned
        
        LoadGrid = False 'set the return value
        m_RowCount = 0 'set the value returned by the "RowCount" property
        
        ListView1(1).ListItems.Add 'Add a blank line
        
        If pNoDataCol = 1 Then
            ListView1(1).ListItems(1).Text = pNoDataMsg 'show the "None" message passed in
        Else
            For i = 1 To pNoDataCol - 2
                'Skip to the desired column to display the "None" message
                ListView1(1).ListItems(1).ListSubItems.Add i
            Next i
            ListView1(1).ListItems(1).ListSubItems.Add pNoDataCol - 1, , pNoDataMsg
        End If
        'Disable the grid
        ListView1(1).Enabled = False
    Else
        'Some records were found
        ListView1(1).Enabled = True
        LoadGrid = True 'set the return value
        
        maxRow = rs.RecordCount
        maxCol = rs.Fields.Count - 1
        
        m_RowCount = maxRow 'set the value returned by the "RowCount" property
        
        For i = 1 To maxRow
            'Add the 1st column data
            ListView1(1).ListItems.Add i, , rs.Fields(0).Value
            
            'Add the rest of the column data
            For curCol = 1 To maxCol
                If Not IsNull(rs.Fields(curCol).Value) Then
                    ListView1(1).ListItems(i).ListSubItems.Add curCol, , rs.Fields(curCol).Value
                Else
                    ListView1(1).ListItems(i).ListSubItems.Add curCol, , ""
                End If
            Next curCol
            
            rs.MoveNext
        Next i
    
    End If
    
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,1,2,0
Public Property Get SelectedCount() As Long
Attribute SelectedCount.VB_Description = "Returns the number of rows that are selected by the user."
Attribute SelectedCount.VB_MemberFlags = "400"

    'Get the number of items in the grid that are selected by the user
    
    Dim i As Long
    
    'start with zero
    m_SelectedCount = 0
    
    If ListView1(1).Checkboxes = True Then
        'We are using checkboxes, so count only those that are checked
        For i = 1 To m_RowCount
            If ListView1(1).ListItems.Item(i).Checked = True Then
                m_SelectedCount = m_SelectedCount + 1
            End If
        Next i
        
    Else 'not using checkboxes
        'We are not using checkboxes, so count those that are "selected"
        For i = 1 To m_RowCount
            If ListView1(1).ListItems.Item(i).Selected = True Then
                m_SelectedCount = m_SelectedCount + 1
            End If
        Next i
    End If
    
    SelectedCount = m_SelectedCount

End Property

Public Property Let SelectedCount(ByVal New_SelectedCount As Long)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_SelectedCount = New_SelectedCount
    PropertyChanged "SelectedCount"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub SetGridData(pRow As Long, pColumn As Long, pData As String)
Attribute SetGridData.VB_Description = "Sets the contents of a cell."
    
    'sets a cell in the grid to a specific value
    'pRow and pColumn start with 1 for editable data
    
    'NOTE: pass zero (0) for pRow if you want to use the currently selected row
    
    Dim rowNum As Long
    
    If pRow = 0 Then
        rowNum = ListView1(1).SelectedItem.Index
    Else
        rowNum = pRow
    End If
        
    If pColumn = 1 Then
        ListView1(1).ListItems.Item(rowNum).Text = pData
    Else
        ListView1(1).ListItems.Item(rowNum).ListSubItems.Item(pColumn - 1).Text = pData
    End If
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function GetGridData(pRow As Long, pColumn As Long) As String
Attribute GetGridData.VB_Description = "Returns the contents of a cell."
    
    'gets a cell in the grid
    'pRow and pColumn start with 1 for editable data
    
    'NOTE: pass zero (0) for pRow if you want to use the currently selected row
    
    Dim rowNum As Long
    
    If pRow = 0 Then
        rowNum = ListView1(1).SelectedItem.Index
    Else
        rowNum = pRow
    End If
        
    If pColumn = 1 Then
        GetGridData = ListView1(1).ListItems.Item(rowNum).Text
    Else
        GetGridData = ListView1(1).ListItems.Item(rowNum).ListSubItems.Item(pColumn - 1).Text
    End If
    
End Function

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Static i As Integer 'This keeps track of the last column that was clicked

    If i = ColumnHeader.Index Then  'clicking a second time
        SortByColumn ColumnHeader.Index, False
        i = -1
    Else
        SortByColumn ColumnHeader.Index, True
        i = ColumnHeader.Index
    End If
    
    RaiseEvent ColumnClick(ColumnHeader)
 
End Sub


Private Sub ListView1_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    RaiseEvent ItemClick(Item)
End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackStyle = m_def_BackStyle
    m_RowCount = m_def_RowCount
    m_SelectedCount = m_def_SelectedCount
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Load ListView1(1)
    ListView1(0).BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    ListView1(0).ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set ListView1(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    ListView1(0).BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    ListView1(0).Checkboxes = PropBag.ReadProperty("Checkboxes", False)
    ListView1(0).MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    m_RowCount = PropBag.ReadProperty("RowCount", m_def_RowCount)
    m_SelectedCount = PropBag.ReadProperty("SelectedCount", m_def_SelectedCount)
    
    ListView1(1).BackColor = ListView1(0).BackColor
    ListView1(1).BorderStyle = ListView1(0).BorderStyle
    ListView1(1).Checkboxes = ListView1(0).Checkboxes
    ListView1(1).Enabled = ListView1(0).Enabled
    ListView1(1).Font = ListView1(0).Font
    ListView1(1).ForeColor = ListView1(0).ForeColor
    ListView1(1).Height = ListView1(0).Height
    ListView1(1).MultiSelect = ListView1(0).MultiSelect
    ListView1(1).TabIndex = ListView1(0).TabIndex
    ListView1(1).TabStop = ListView1(0).TabStop
    ListView1(1).Tag = ListView1(0).Tag
    ListView1(1).Visible = True
    ListView1(1).Width = ListView1(0).Width

End Sub

Private Sub UserControl_Resize()

On Error Resume Next
    'This is the code that allows the control to be resized at design time
    ListView1(1).Height = UserControl.Height
    ListView1(1).Width = UserControl.Width
    
End Sub

Private Sub UserControl_Show()
    
    'Set up the tooltips
    Set TT = New CTooltip
    TT.Style = TTBalloon
    TT.Icon = TTNoIcon

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("BackColor", ListView1(1).BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", ListView1(1).ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", ListView1(1).Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", ListView1(1).BorderStyle, 1)
    Call PropBag.WriteProperty("Checkboxes", ListView1(1).Checkboxes, False)
    Call PropBag.WriteProperty("MultiSelect", ListView1(1).MultiSelect, False)
    Call PropBag.WriteProperty("RowCount", m_RowCount, m_def_RowCount)
    Call PropBag.WriteProperty("SelectedCount", m_SelectedCount, m_def_SelectedCount)
    
End Sub


