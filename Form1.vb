Imports System.Collections.ObjectModel
Imports System.IO
Imports System.Collections

Public Structure tList
	Public Name As String
	Public AltName As String
	Public idx As Short
End Structure

Public Class Form1
	Const CrLf As String = Chr(13) + Chr(10)

	Dim books As tList
	Dim CfgFile, ExePath, FilePath, WorkPath, FileOut, Avoid(77), DirPath(77), DrvType(77), Exclude(77), UpDir As String
	Dim Bt As Byte
	Dim B2 As Char
	Dim OldEntry, NewEntry, FULLtext, FinalText, SearchText, SearchTextCap, Ttext, Ntext, FilterText As String
	Dim Avoids, Count, Ct, Db, DirCount, Excludes, MaxEntries, DirSize(77) As Integer
	Dim lvItem As ListViewItem
	'Dim HasWeapon As SByte
	'Dim Multiplier As Single
	'Dim QQ As Long

	'ToDo:  [Save Info] button needs updating

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim A As Integer
		'Dim gvc1 As New GridViewColumn()
		'gvc1.DisplayMemberBinding = New Binding("FirstName")
		'gvc1.Header = "FirstName"
		'gvc1.Width = 100
		'ListView1.Columns.Clear()
		'ListView1.Columns.Add("Size")
		'ListView1.Columns.Add("Path")
		'Explorer1.Show()
		ListView1.HeaderStyle = ColumnHeaderStyle.Clickable

		UpDir = " << Previous Directory"
		ExePath = My.Computer.FileSystem.CurrentDirectory + "\"
		CfgFile = ExePath + "Dir Info.ini"
		FileSystem.FileOpen(1, CfgFile, OpenMode.Append, OpenAccess.Write)
		FileSystem.FileClose(1)
		Dim columnHeader1 As New ColumnHeader
		With columnHeader1
            .Text = "Path"
            .TextAlign = HorizontalAlignment.Center
            .Width = 300
        End With
		Dim columnHeader2 As New ColumnHeader
		With columnHeader2
            .Text = "Info"
            '.TextAlign = HorizontalAlignment.Center
            .TextAlign = HorizontalAlignment.Right
            .Width = 70
        End With
		Me.ListView1.Columns.Add(columnHeader1)
        Me.ListView1.Columns.Add(columnHeader2)

        Dim inMode As Integer
		FileSystem.FileOpen(1, CfgFile, OpenMode.Input, OpenAccess.Read)
		For A = 1 To 70
			If EOF(1) = False Then
				FileSystem.Input(1, FULLtext)
				FULLtext = LTrim(FULLtext)
				FULLtext = RTrim(FULLtext)
				Ttext = LowerCase(FULLtext)
				If Ttext = "[recent]" Then inMode = 1
				If Ttext = "[exclude]" Then inMode = 2
				If Ttext = "[avoid]" Then inMode = 3
				If InStr(Ttext, "[") = 0 And Ttext > "" Then
					Select Case inMode
						Case 1
							WorkPath = FULLtext
						Case 2
							Excludes += 1
							Exclude(Excludes) = FULLtext
						Case 3
							Avoids += 1
							Avoid(Avoids) = FULLtext
					End Select
				End If
			End If
		Next
		FileSystem.FileClose(1)
		If Len(WorkPath) < 3 Then WorkPath = Mid(ExePath, 1, 3)
		SaveCfg()

		DrvType(2) = "Removable"
		DrvType(3) = "Hard Drive"
		DrvType(5) = "CD or DVD"
		FastLoad()
	End Sub

	Private Sub SaveCfg()
		Dim A As Integer
		FileSystem.FileOpen(1, CfgFile, OpenMode.Output, OpenAccess.Write)
		Ttext = "[Recent]" + CrLf
		FileSystem.Print(1, Ttext)
		Ttext = WorkPath + CrLf
		FileSystem.Print(1, Ttext)
		FileSystem.Print(1, CrLf)

		Ttext = "[Exclude]" + CrLf
		FileSystem.Print(1, Ttext)
		For A = 1 To Excludes
			Ttext = Exclude(A) + CrLf
			FileSystem.Print(1, Ttext)
		Next A
		FileSystem.Print(1, CrLf)

		Ttext = "[avoid]" + CrLf
		FileSystem.Print(1, Ttext)
		For A = 1 To Avoids
			Ttext = Avoid(A) + CrLf
			FileSystem.Print(1, Ttext)
		Next A
		FileSystem.FileClose(1)
	End Sub

	Private Function LowerCase(ByVal Text1 As String) As String
		Dim Text2 As String
		Dim Pos, W As Short
		For W = 1 To Len(Text1)
			Text2 = Mid(Text1, W, 1)
			Pos = InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Text2)
			If Pos Then
				Text2 = Mid("abcdefghijklmnopqrstuvwxyz", Pos, 1)
			End If
			Mid(Text1, W, 1) = Text2
		Next
		Return Text1
	End Function

	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
		'~~~~~  Evaluate  ~~~~~
		Loadup()
	End Sub

	Private Sub Loadup()
		Dim A, B As Integer
		Dim Skip As Boolean
		TextBox2.Text = WorkPath
		ListView1.Items.Clear()
		Dim Temp1, Temp2 As Int64
		Dim dirs1 As ReadOnlyCollection(Of String)
		dirs1 = My.Computer.FileSystem.GetDirectories(WorkPath, FileIO.SearchOption.SearchTopLevelOnly)
		DirCount = dirs1.Count

		ReDim DirPath(DirCount + 1), DirSize(DirCount + 1)
		ProgressBar1.Maximum = DirCount
		For A = 1 To DirCount
			FULLtext = dirs1.Item(A - 1)
			Ttext = Replace(FULLtext, WorkPath, "", , , CompareMethod.Text)
			Skip = False
			For B = 1 To Excludes
				If Ttext = Exclude(B) Then Skip = True
			Next
			If Skip = False Then
				lvItem = ListView1.Items.Add(Ttext)
				For B = 1 To Avoids
					If Ttext = Avoid(B) Then Skip = True
				Next
				If Skip Then
					Temp1 = -1
				Else
					Temp1 = FindDirSize(FULLtext)
					Temp2 = 0
				End If
				Ttext = ""
				If Temp1 < 1000 Then
					Ntext = LTrim(Str(Temp1)) + "  B"
					If Temp1 < 0 Then Ntext = "Avoided"
				Else
					Temp1 = Temp1 / 1000
					Ntext = LTrim(Str(Temp1))
					For B = 1 To Len(Ntext)
						Temp2 = 1 + Len(Ntext) - B
						Select Case Temp2
							Case 3, 6, 9, 12, 15
								If Len(Ntext) > Temp2 Then
									Ttext = Ttext + ","
								End If
							Case Else
						End Select
						Ttext = Ttext + Mid(Ntext, B, 1)
					Next
					Ntext = Ttext + "  K"
				End If
				If Len(Ntext) < 10 Then
					Ntext = "          " + Ntext
					Ntext = RightStr(Ntext, 10)
				End If
			Else
				Ntext = "Unknown"
			End If
			'DirSize(A) = Temp
			lvItem.SubItems.AddRange(New String() {Ntext})
			ProgressBar1.Value = A
		Next
		AdjustListWidth()
		SaveCfg()
	End Sub

	Private Sub FastLoad()
		WorkPath = Replace(WorkPath, "\\", "\")
		Dim A, B As Integer
		Dim Skip As Boolean
		TextBox2.Text = WorkPath
		'ListBox1.Items.Clear()
		ListView1.Items.Clear()
		Dim dirs1 As ReadOnlyCollection(Of String)
		dirs1 = My.Computer.FileSystem.GetDirectories(WorkPath, FileIO.SearchOption.SearchTopLevelOnly)
		DirCount = dirs1.Count
		ReDim DirPath(DirCount + 1)

		ProgressBar1.Value = 0
		'ListBox1.Items.Add(UpDir)
		Ntext = " "
		For A = 1 To DirCount
			Ttext = dirs1.Item(A - 1)
			Ttext = Replace(Ttext, WorkPath, "", , , CompareMethod.Text)
			Skip = False
			For B = 1 To Excludes
				If Ttext = Exclude(B) Then Skip = True
			Next
			If Skip = False Then
				Dim listItem As New ListViewItem(Ttext)
				listItem.SubItems.Add(Ntext)
				ListView1.Items.Add(listItem)
				'ListBox1.Items.Add(Ttext)
			End If
			'ListView1.Items.Add(Ttext)
			'ListView1.Items.Add(ttext).SubItems.
			'ListView1.Items(0).SubItems.Item(B).Text = Ttext
			'ListView1.Items(1).SubItems.Add(Ntext)
			'lvItem.SubItems.AddRange(New String() {Ntext})
		Next
		AdjustListWidth()
		SaveCfg()
	End Sub

	Private Sub ListDrives()
        Dim A As Integer
        TextBox2.Text = "This Computer"
		'ListBox1.Items.Clear()
		ListView1.Items.Clear()
		'Dim Temp1 As Integer
		Dim allDrives() As DriveInfo = DriveInfo.GetDrives()
		Dim d As DriveInfo

		For Each d In allDrives
			Ttext = RTrim(d.Name)
			Ntext = DrvType(d.DriveType)
			'ListBox1.Items.Add(Ttext)
			'ListView1.Items.Add(Ttext)
			lvItem = ListView1.Items.Add(Ttext)
			lvItem.SubItems.AddRange(New String() {Ntext})
			'Console.WriteLine("  File type: {0}", d.DriveType)
			If True = False Then 'd.IsReady = True Then
				Console.WriteLine("  Volume label: {0}", d.VolumeLabel)
				Console.WriteLine("  File system: {0}", d.DriveFormat)
				Console.WriteLine("  Available space to current user:{0, 15} bytes", d.AvailableFreeSpace)
				Console.WriteLine("  Total available space:          {0, 15} bytes", d.TotalFreeSpace)
				Console.WriteLine("  Total size of drive:            {0, 15} bytes ", d.TotalSize)
			End If
		Next

		'Dim drives1 As System.IO.DriveInfo()
		'drives1 = System.IO.DriveInfo.GetDrives
		'DirCount = drives1.Count
		'DirCount = drives1.GetValue
		'ReDim DirPath(DirCount + 1)
		ProgressBar1.Value = 0

		For A = 1 To DirCount
			'Ttext = drives1.ElementAt(A - 1).Name
			'Temp1 = drives1.ElementAt(A - 1).DriveType
			'Ttext = Ttext + "    " + DrvType(Temp1)
			'ListBox1.Items.Add(Ttext)
		Next
		AdjustListWidth()
	End Sub

	Private Function FindDirSize(ByVal Path1 As String)
        Dim A As Integer
        Dim Files1 As ReadOnlyCollection(Of String)
		Dim Info1 As System.IO.FileInfo
		Dim FileCount1 As Integer
		Dim TotalSize1 As Int64
		Dim Text1 As String
		'Dim Filepath1(70) As String

		Files1 = My.Computer.FileSystem.GetFiles(Path1, FileIO.SearchOption.SearchAllSubDirectories)
		FileCount1 = Files1.Count
		'ReDim Filepath1(FileCount1)
		TotalSize1 = 0
		For A = 1 To FileCount1
			Text1 = Files1.Item(A - 1)
			'Ttext = Replace(Ttext, WorkPath, "", , , CompareMethod.Text)
			Info1 = My.Computer.FileSystem.GetFileInfo(Text1)
			TotalSize1 = TotalSize1 + Info1.Length
		Next
		Return TotalSize1
	End Function

	Private Function RightStr(ByVal Text1 As String, ByVal Pos1 As Integer)
		Dim Len1 As Integer
		Len1 = Len(Text1)
		If Pos1 > Len1 Then Pos1 = Len1
		Text1 = Mid(Text1, 1 + Len1 - Pos1, Pos1)
		Return Text1
	End Function

	Private Sub NoLongerUsed()
		'ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 
		'Handles ListBox1.SelectedIndexChanged
		Dim Text1 As String
		Dim idx1, Len1, Pos1 As Integer

		idx1 = ListBox1.SelectedIndex
		If idx1 < 0 Then idx1 = 0
		Text1 = ListBox1.Items.Item(idx1)
		Pos1 = InStr(Text1, " K ")
		If Pos1 = 0 Then
			If Len(Text1) < 4 And InStr(Text1, ":") Then
				WorkPath = Text1
				FastLoad()
			Else
				WorkPath = WorkPath + Text1 + "\"
				FastLoad()
			End If
		Else
			Len1 = Len(Text1)
			'Text1 = Mid(Text1, Pos1 + 2, Len1)
			Mid(Text1, 1, Pos1 + 2) = Mid("             ", 1, Pos1 + 2)
			WorkPath = WorkPath + LTrim(Text1) + "\"
			FastLoad()
		End If
		'MsgBox(File1)
		'Dim FileCount1, TotalSize1 As Integer
		'Dim Text1 As String
	End Sub

	Private Sub GoUpDir()
		Dim Len1, Pos1 As Integer
		Len1 = Len(WorkPath)
		Pos1 = RevStr(2, WorkPath, "\")
		If Pos1 Then
			WorkPath = Mid(WorkPath, 1, Pos1)
			FastLoad()
		Else
			WorkPath = ""
			ListDrives()
		End If
	End Sub

	Private Function RevStr(ByVal Start1 As Integer, ByVal Text1 As String, ByVal Text2 As String)
		Dim A, B As Integer
		Dim Count1, End1, Len2 As Integer
		If Start1 < 1 Then Start1 = 1
		Count1 = Len(Text1) + 1 - Start1
		Len2 = Len(Text2)
		For A = Count1 To 1 Step -1
			If Mid(Text1, A, Len2) = Text2 Then
				End1 = A
				A = 0
			End If
		Next
		Return End1
	End Function

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
		GoUpDir()
	End Sub

	Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
		Dim A, B As Integer
		Ttext = Replace(WorkPath, ":", "")
		Ttext = Replace(Ttext, "\", ".")
		FileOut = ExePath + Ttext + "txt"
		FileSystem.FileOpen(1, FileOut, OpenMode.Output, OpenAccess.Write)
		For A = 1 To DirCount
			'Ttext = ListBox1.Items.Item(A - 1) + Chr(13) + Chr(10)
			FileSystem.Print(1, Ttext)
		Next
		FileSystem.FileClose(1)
	End Sub

	Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
		If TextBox1.Text > "" Then BeginSearch()
	End Sub

	Private Sub BeginSearch()
		Dim A, B, Pos1 As Integer
		Dim SearchText As String
		SearchText = TextBox1.Text
		ListView1.Items.Clear()
		Dim Files1 As ReadOnlyCollection(Of String)
		Files1 = My.Computer.FileSystem.GetFiles(WorkPath, FileIO.SearchOption.SearchAllSubDirectories)
		DirCount = Files1.Count

		ReDim DirPath(DirCount + 1), DirSize(DirCount + 1)
		ProgressBar1.Maximum = DirCount
		'ListBox1.Items.Add(UpDir)
		For A = 1 To DirCount
			FULLtext = Files1.Item(A - 1)
			Pos1 = RevStr(2, FULLtext, "\")
			If Pos1 Then
				Ttext = Mid(FULLtext, 1, Pos1)
				Pos1 = Len(FULLtext) - Pos1
				Ntext = RightStr(FULLtext, Pos1)
			Else
				Ttext = FULLtext
				Ntext = ""
			End If
			If InStr(Ntext, SearchText, CompareMethod.Text) Then
				lvItem = ListView1.Items.Add(Ttext)
				lvItem.SubItems.AddRange(New String() {Ntext})
			End If
			ProgressBar1.Value = A
		Next
		AdjustListWidth()
	End Sub

	Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
		Dim Text1 As String
		Dim A, B, idx1, Pos1 As Integer
		Dim Skip As Boolean
		Dim breakfast As ListView.SelectedListViewItemCollection = Me.ListView1.SelectedItems
		Dim item As ListViewItem
		Dim price As Double = 0.0
		For Each item In breakfast
			Text1 = item.SubItems(0).Text
		Next
		idx1 = ListView1.FocusedItem.Index
		'   x +=y
		'TextBox1.Text = CType(price, String)
		' idx1 = ListView1.ListItems.Item(I)
		' idx1 = breakfast.ListItems.Count
		'idx1 = ListView1.SelectedIndices()
		'idx1 = ListView1.SelectedItems(0).Index
		'idx1 = ListView1.SelectedIndices.Count >0
		If idx1 < 0 Then idx1 = 0
		Text1 = ListView1.FocusedItem.Text
		'Items.Item(idx1)
		'Pos1 = InStr(Text1, " K ")
		'If Pos1 = 0 Then
		If InStr(Text1, ":") Then 'Len(Text1) < 4 And
			WorkPath = Text1
			FastLoad()
		Else
			Skip = False
			For B = 1 To Excludes
				If Ttext = Exclude(B) Then Skip = True
			Next
			If Skip = False Then
				WorkPath = WorkPath + Text1 + "\"
				FastLoad()
			End If
		End If
	End Sub

	Private Sub AdjustListWidth()
		ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
		Dim HalfWidth, Width1, Width2 As Integer
		HalfWidth = ListView1.Width / 2
		Width1 = ListView1.Columns(0).Width
		Width2 = ListView1.Columns(1).Width
		If Width1 < HalfWidth Then Width1 = HalfWidth
		If Width2 < 77 Then Width2 = 77
		'If Width1 + Width2 < 374 Then Width2 = 374 - Width1
		ListView1.Columns(0).Width = Width1
		ListView1.Columns(1).Width = Width2
	End Sub
End Class
