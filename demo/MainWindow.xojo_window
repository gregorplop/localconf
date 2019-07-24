#tag Window
Begin Window MainWindow
   BackColor       =   &cFFFFFF00
   Backdrop        =   0
   CloseButton     =   True
   Compatibility   =   ""
   Composite       =   False
   Frame           =   0
   FullScreen      =   False
   FullScreenButton=   False
   HasBackColor    =   False
   Height          =   594
   ImplicitInstance=   True
   LiveResize      =   True
   MacProcID       =   0
   MaxHeight       =   32000
   MaximizeButton  =   True
   MaxWidth        =   32000
   MenuBar         =   0
   MenuBarVisible  =   True
   MinHeight       =   594
   MinimizeButton  =   True
   MinWidth        =   1152
   Placement       =   0
   Resizeable      =   True
   Title           =   "localconf_demo"
   Visible         =   True
   Width           =   1152
   Begin TextField localconf_filenameField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   0
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   19.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   674
   End
   Begin PushButton openBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Open"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   960
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   1
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin PushButton closeBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Close"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   1052
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   2
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin Listbox dumpList
      AutoDeactivate  =   True
      AutoHideScrollbars=   True
      Bold            =   False
      Border          =   True
      ColumnCount     =   1
      ColumnsResizable=   False
      ColumnWidths    =   ""
      DataField       =   ""
      DataSource      =   ""
      DefaultRowHeight=   -1
      Enabled         =   True
      EnableDrag      =   False
      EnableDragReorder=   False
      GridLinesHorizontal=   0
      GridLinesVertical=   0
      HasHeading      =   False
      HeadingIndex    =   -1
      Height          =   344
      HelpTag         =   ""
      Hierarchical    =   False
      Index           =   -2147483648
      InitialParent   =   ""
      InitialValue    =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      RequiresSelection=   False
      Scope           =   0
      ScrollbarHorizontal=   False
      ScrollBarVertical=   True
      SelectionType   =   0
      ShowDropIndicator=   False
      TabIndex        =   3
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   114
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   1112
      _ScrollOffset   =   0
      _ScrollWidth    =   -1
   End
   Begin Label Label1
      AutoDeactivate  =   True
      Bold            =   True
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   True
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   7
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "Database dump"
      TextAlign       =   0
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   67
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   140
   End
   Begin PushButton refreshBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Query"
      Default         =   True
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   868
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   9
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   67
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin TextField passwordField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   ""
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   706
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   True
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   14
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   19.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin PushButton CreateBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Create"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   868
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   18
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   20
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin TextField applicationField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "application"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "application"
      Index           =   -2147483648
      Italic          =   False
      Left            =   129
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   19
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin TextField userField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "user"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "user"
      Index           =   -2147483648
      Italic          =   False
      Left            =   291
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   20
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin TextField sectionField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "section"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "section"
      Index           =   -2147483648
      Italic          =   False
      Left            =   453
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   21
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin TextField keyField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "key"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "key"
      Index           =   -2147483648
      Italic          =   False
      Left            =   615
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   22
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin TextField commentField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFE1E100
      Bold            =   False
      Border          =   True
      CueText         =   "comment"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "comment"
      Index           =   -2147483648
      Italic          =   False
      Left            =   982
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   23
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin TextField valueField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFD900
      Bold            =   False
      Border          =   True
      CueText         =   "value"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "value"
      Index           =   -2147483648
      Italic          =   False
      Left            =   784
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   24
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   150
   End
   Begin PushButton UpsertBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Upsert"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   982
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   25
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   512
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   150
   End
   Begin TextField WHEREfield
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "WHERE clauses or objidx (for Upsert, Delete, Read Single)"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      Italic          =   False
      Left            =   172
      LimitText       =   0
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   True
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   26
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   19.0
      TextUnit        =   0
      Top             =   67
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   684
   End
   Begin PushButton queryDistinctBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Distinct"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   960
      LockBottom      =   False
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   True
      Scope           =   0
      TabIndex        =   27
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   67
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   80
   End
   Begin PushButton appendBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Append"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   784
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Scope           =   0
      TabIndex        =   28
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   512
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   150
   End
   Begin Label MainLabel
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   20
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   True
      LockTop         =   False
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   29
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "localconf demo application"
      TextAlign       =   1
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   559
      Transparent     =   True
      Underline       =   False
      Visible         =   True
      Width           =   788
   End
   Begin PushButton DeleteBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Delete"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   615
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Scope           =   0
      TabIndex        =   30
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   512
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   150
   End
   Begin PushButton ReadSingleBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Read Single"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   129
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Scope           =   0
      TabIndex        =   31
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   512
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   150
   End
   Begin PushButton ReadArrayBtn
      AutoDeactivate  =   True
      Bold            =   False
      ButtonStyle     =   "0"
      Cancel          =   False
      Caption         =   "Read Array"
      Default         =   False
      Enabled         =   True
      Height          =   35
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   291
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Scope           =   0
      TabIndex        =   32
      TabPanelIndex   =   0
      TabStop         =   True
      TextFont        =   "System"
      TextSize        =   14.0
      TextUnit        =   0
      Top             =   512
      Transparent     =   False
      Underline       =   False
      Visible         =   True
      Width           =   150
   End
   Begin Label projectLink
      AutoDeactivate  =   True
      Bold            =   False
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Height          =   25
      HelpTag         =   ""
      Index           =   -2147483648
      InitialParent   =   ""
      Italic          =   False
      Left            =   820
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   False
      LockRight       =   True
      LockTop         =   False
      Multiline       =   False
      Scope           =   0
      Selectable      =   False
      TabIndex        =   33
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   "localconf on GitHub"
      TextAlign       =   2
      TextColor       =   &c0000FF00
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   564
      Transparent     =   True
      Underline       =   True
      Visible         =   True
      Width           =   312
   End
   Begin TextField languageField
      AcceptTabs      =   False
      Alignment       =   0
      AutoDeactivate  =   True
      AutomaticallyCheckSpelling=   False
      BackColor       =   &cFFFFFF00
      Bold            =   False
      Border          =   True
      CueText         =   "language"
      DataField       =   ""
      DataSource      =   ""
      Enabled         =   True
      Format          =   ""
      Height          =   30
      HelpTag         =   "language"
      Index           =   -2147483648
      Italic          =   False
      Left            =   20
      LimitText       =   0
      LockBottom      =   True
      LockedInPosition=   False
      LockLeft        =   True
      LockRight       =   False
      LockTop         =   False
      Mask            =   ""
      Password        =   False
      ReadOnly        =   False
      Scope           =   0
      TabIndex        =   34
      TabPanelIndex   =   0
      TabStop         =   True
      Text            =   ""
      TextColor       =   &c00000000
      TextFont        =   "System"
      TextSize        =   16.0
      TextUnit        =   0
      Top             =   470
      Transparent     =   False
      Underline       =   False
      UseFocusRing    =   True
      Visible         =   True
      Width           =   97
   End
End
#tag EndWindow

#tag WindowCode
	#tag Event
		Sub Open()
		  UI_mode("CLOSED")
		  
		  
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Function findColumnIdx(columnName as string) As Integer
		  for i as Integer = 0 to dumpList.ColumnCount - 1
		    if columnName = dumpList.Heading(i) then Return i
		  next i
		  
		  Return -1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub query()
		  dumpList.DeleteAllRows
		  
		  dim allRecords(-1) as localconfRecord = myLocalconf.QueryGeneric(WHEREfield.Text.Trim)
		  
		  if allRecords.Ubound < 0 then
		    MsgBox "No configuration entries!"
		    return
		  end if
		  
		  if allRecords.Ubound = 0 and allRecords(0).Error then
		    MsgBox allRecords(0).ErrorMessage
		    return
		  end if
		  
		  showRecords(allRecords)
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub showRecords(records() as localconfRecord)
		  dumpList.DeleteAllRows
		  
		  dim row(7) as String
		  
		  for i as Integer = 0 to records.Ubound
		    
		    row(0) = str(records(i).objidx)
		    row(1) = records(i).language
		    row(2) = records(i).application
		    row(3) = records(i).user
		    row(4) = records(i).section
		    row(5) = records(i).key
		    row(6) = records(i).value
		    row(7) = records(i).comment
		    
		    dumpList.AddRow row
		    
		  next i
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub UI_mode(mode as string)
		  select case mode.Uppercase
		    
		  case "OPEN"
		    
		    localconf_filenameField.Enabled = false
		    passwordField.Enabled = False
		    openBtn.Enabled = false
		    closeBtn.Enabled = true
		    CreateBtn.Enabled = False
		    
		    WHEREfield.Enabled = true
		    
		    UpsertBtn.Enabled = true
		    
		    dumpList.Enabled = true
		    dumpList.DeleteAllRows
		    refreshBtn.Enabled = true
		    queryDistinctBtn.Enabled = true
		    DeleteBtn.Enabled = true
		    appendBtn.Enabled = true
		    ReadArrayBtn.Enabled = true
		    ReadSingleBtn.Enabled = true
		    
		    Label1.Enabled = true
		    
		    
		  case "CLOSED"
		    
		    localconf_filenameField.Enabled = true
		    passwordField.Enabled = true
		    openBtn.Enabled = true
		    closeBtn.Enabled = False
		    CreateBtn.Enabled = true
		    
		    WHEREfield.Enabled = false
		    WHEREfield.Text = ""
		    
		    UpsertBtn.Enabled = false
		    
		    dumpList.Enabled = False
		    dumpList.DeleteAllRows
		    refreshBtn.Enabled = False
		    queryDistinctBtn.Enabled = false
		    appendBtn.Enabled = false
		    DeleteBtn.Enabled = false
		    ReadArrayBtn.Enabled = false
		    ReadSingleBtn.Enabled = false
		    
		    Label1.Enabled = false
		    
		  end select
		  
		End Sub
	#tag EndMethod


	#tag Property, Flags = &h0
		myLocalconf As localconfSession
	#tag EndProperty


#tag EndWindowCode

#tag Events localconf_filenameField
	#tag Event
		Sub Open()
		  me.Text = SpecialFolder.Desktop.Child("demo.localconf").NativePath
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events openBtn
	#tag Event
		Sub Action()
		  myLocalconf = new localconfSession(GetFolderItem(localconf_filenameField.Text) , passwordField.Text)
		  
		  if myLocalconf.LastError = "" then
		    UI_mode("OPEN")
		    
		    WHEREfield.Text = ""
		    query
		    
		    
		  else
		    MsgBox myLocalconf.LastError
		    myLocalconf = nil
		  end if
		  
		  
		  
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events closeBtn
	#tag Event
		Sub Action()
		  myLocalconf = nil
		  
		  UI_mode("CLOSED")
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events dumpList
	#tag Event
		Sub Open()
		  me.ColumnCount = 8
		  
		  me.Heading(0) = "objidx"
		  me.Heading(1) = "language"
		  me.Heading(2) = "application"
		  me.Heading(3) = "user"
		  me.Heading(4) = "section"
		  me.Heading(5) = "key"
		  me.Heading(6) = "value"
		  me.Heading(7) = "comment"
		  
		  me.HasHeading = true
		  me.HeaderType(-1) = Listbox.HeaderTypes.NotSortable
		  me.ColumnsResizable = true
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub DoubleClick()
		  Dim row, column As Integer
		  row = Me.RowFromXY(System.MouseX - Me.Left - Self.Left, System.MouseY - Me.Top - Self.Top)
		  
		  if row >= 0 then
		    languageField.Text = me.cell(row,1)
		    applicationField.Text = me.cell(row,2)
		    userField.Text = me.cell(row,3)
		    sectionField.Text = me.cell(row,4)
		    keyField.Text = me.cell(row,5)
		    valueField.Text = me.cell(row,6)
		    commentField.Text = me.cell(row,7)
		    
		  end if
		  
		End Sub
	#tag EndEvent
	#tag Event
		Sub Change()
		  if me.ListIndex < 0 then
		    languageField.Text = ""
		    applicationField.Text = ""
		    userField.Text = ""
		    sectionField.Text = ""
		    keyField.Text = ""
		    valueField.Text = ""
		    commentField.Text = ""
		    
		  end if
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events refreshBtn
	#tag Event
		Sub Action()
		  if IsNull(myLocalconf) = false then query
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events CreateBtn
	#tag Event
		Sub Action()
		  dim createOutcome as String = localconfSession.create(GetFolderItem(localconf_filenameField.Text) , passwordField.Text)
		  
		  if createOutcome = "" then
		    MsgBox "localconf file created OK"
		  else
		    MsgBox "Error: " + EndOfLine + createOutcome
		  end if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events UpsertBtn
	#tag Event
		Sub Action()
		  dim confObj as new localconfRecord(true)
		  
		  if IsNumeric(WHEREfield.Text) then
		    
		    confObj.objidx = WHEREfield.Text.Trim.Val
		    
		  else
		    
		    confObj.application = applicationField.Text.Trim
		    confObj.user = userField.Text.Trim
		    confObj.section = sectionField.Text.Trim
		    confObj.key = keyField.Text.Trim
		    
		  end if
		  
		  confObj.value = valueField.Text.Trim
		  confObj.comment = commentField.Text.Trim
		  
		  confObj = myLocalconf.Upsert(confObj)
		  
		  if confObj.Error then
		    MsgBox confObj.ErrorMessage
		  end if
		  
		  query
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events queryDistinctBtn
	#tag Event
		Sub Action()
		  if WHEREfield.Text.CountFields(",") <> 2 then 
		    MsgBox "Input parameters format should be <field>,<where>"
		    Return
		  end if
		  
		  dim field as String = WHEREfield.Text.NthField("," , 1)
		  dim where as String = WHEREfield.Text.NthField("," , 2)
		  
		  dim fields(-1) as String = myLocalconf.QueryDistinct(field , where)
		  
		  if myLocalconf.LastError <> "" then 
		    MsgBox "Error: " + EndOfLine + myLocalconf.LastError
		    
		  else
		    
		    MsgBox Join(fields , EndOfLine)
		    
		  end if
		  
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events appendBtn
	#tag Event
		Sub Action()
		  dim confObj as new localconfRecord(true)
		  
		  confObj.application = applicationField.Text.Trim
		  confObj.user = userField.Text.Trim
		  confObj.section = sectionField.Text.Trim
		  confObj.key = keyField.Text.Trim
		  confObj.value = valueField.Text.Trim
		  confObj.comment = commentField.Text.Trim
		  
		  confObj = myLocalconf.Append2Array(confObj)
		  
		  if confObj.Error then
		    MsgBox confObj.ErrorMessage
		  else
		    MainLabel.text = "Just added objidx = " + str(confObj.objidx)
		  end if
		  
		  query
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events DeleteBtn
	#tag Event
		Sub Action()
		  dim confObj as new localconfRecord(true)
		  
		  if IsNumeric(WHEREfield.Text) then
		    
		    confObj.objidx = WHEREfield.Text.Trim.Val
		    
		  else
		    
		    confObj.application = applicationField.Text.Trim
		    confObj.user = userField.Text.Trim
		    confObj.section = sectionField.Text.Trim
		    confObj.key = keyField.Text.Trim
		    
		  end if
		  
		  confObj = myLocalconf.Delete(confObj)
		  
		  if confObj.Error then
		    MsgBox confObj.ErrorMessage
		  else
		    MainLabel.text = "Deleted record"
		  end if
		  
		  query
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ReadSingleBtn
	#tag Event
		Sub Action()
		  dim confObj as new localconfRecord(true)
		  
		  if IsNumeric(WHEREfield.Text) then
		    
		    confObj.objidx = WHEREfield.Text.Trim.Val
		    
		  else
		    
		    confObj.language = languageField.Text.Trim
		    confObj.application = applicationField.Text.Trim
		    confObj.user = userField.Text.Trim
		    confObj.section = sectionField.Text.Trim
		    confObj.key = keyField.Text.Trim
		    
		  end if
		  
		  confObj = myLocalconf.ReadSingle(confObj)
		  
		  if confObj.Error then
		    MsgBox confObj.ErrorMessage
		  else
		    MainLabel.text = "Read record"
		    MsgBox "Value: " + confObj.value + EndOfLine + "Comment: " + confObj.comment
		  end if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events ReadArrayBtn
	#tag Event
		Sub Action()
		  dim confObj as new localconfRecord(true)
		  
		  confObj.application = applicationField.Text.Trim
		  confObj.user = userField.Text.Trim
		  confObj.section = sectionField.Text.Trim
		  confObj.key = keyField.Text.Trim
		  
		  dim confArray(-1) as localconfRecord = myLocalconf.ReadArray(confObj)
		  
		  if confArray(0).Error then
		    MainLabel.Text = "Error getting array"
		    MsgBox confArray(0).ErrorMessage
		  ElseIf confArray(0).Exists = false then
		    MainLabel.Text = "No elements in array"
		  else
		    
		    dim msg as String
		    
		    for i as Integer = 0 to confArray.Ubound
		      msg = msg + "Value: " + confArray(i).value + EndOfLine
		      msg = msg + "Comment: " + confArray(i).comment + EndOfLine
		      msg = msg + "-------------"+ EndOfLine
		    next i
		    
		    MsgBox msg
		    
		  end if
		  
		End Sub
	#tag EndEvent
#tag EndEvents
#tag Events projectLink
	#tag Event
		Function MouseDown(X As Integer, Y As Integer) As Boolean
		  ShowURL(localconfSession.projectURL)
		  
		End Function
	#tag EndEvent
#tag EndEvents
#tag ViewBehavior
	#tag ViewProperty
		Name="Name"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Interfaces"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Super"
		Visible=true
		Group="ID"
		Type="String"
		EditorType="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Width"
		Visible=true
		Group="Size"
		InitialValue="600"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Height"
		Visible=true
		Group="Size"
		InitialValue="400"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinWidth"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinHeight"
		Visible=true
		Group="Size"
		InitialValue="64"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxWidth"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaxHeight"
		Visible=true
		Group="Size"
		InitialValue="32000"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Frame"
		Visible=true
		Group="Frame"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Document"
			"1 - Movable Modal"
			"2 - Modal Dialog"
			"3 - Floating Window"
			"4 - Plain Box"
			"5 - Shadowed Box"
			"6 - Rounded Window"
			"7 - Global Floating Window"
			"8 - Sheet Window"
			"9 - Metal Window"
			"11 - Modeless Dialog"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Title"
		Visible=true
		Group="Frame"
		InitialValue="Untitled"
		Type="String"
	#tag EndViewProperty
	#tag ViewProperty
		Name="CloseButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Resizeable"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MaximizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MinimizeButton"
		Visible=true
		Group="Frame"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreenButton"
		Visible=true
		Group="Frame"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Composite"
		Group="OS X (Carbon)"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MacProcID"
		Group="OS X (Carbon)"
		InitialValue="0"
		Type="Integer"
	#tag EndViewProperty
	#tag ViewProperty
		Name="FullScreen"
		Group="Behavior"
		InitialValue="False"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="ImplicitInstance"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="LiveResize"
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Placement"
		Visible=true
		Group="Behavior"
		InitialValue="0"
		Type="Integer"
		EditorType="Enum"
		#tag EnumValues
			"0 - Default"
			"1 - Parent Window"
			"2 - Main Screen"
			"3 - Parent Window Screen"
			"4 - Stagger"
		#tag EndEnumValues
	#tag EndViewProperty
	#tag ViewProperty
		Name="Visible"
		Visible=true
		Group="Behavior"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="HasBackColor"
		Visible=true
		Group="Background"
		InitialValue="False"
		Type="Boolean"
	#tag EndViewProperty
	#tag ViewProperty
		Name="BackColor"
		Visible=true
		Group="Background"
		InitialValue="&hFFFFFF"
		Type="Color"
	#tag EndViewProperty
	#tag ViewProperty
		Name="Backdrop"
		Visible=true
		Group="Background"
		Type="Picture"
		EditorType="Picture"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBar"
		Visible=true
		Group="Menus"
		Type="MenuBar"
		EditorType="MenuBar"
	#tag EndViewProperty
	#tag ViewProperty
		Name="MenuBarVisible"
		Visible=true
		Group="Deprecated"
		InitialValue="True"
		Type="Boolean"
		EditorType="Boolean"
	#tag EndViewProperty
#tag EndViewBehavior
