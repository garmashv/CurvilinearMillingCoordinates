REM ---------------------------- Module1 - 1 ----------------------------
Public filename As String 
Public setki As New Collection 
Public setka As AcadPolygonMesh 
Public coor As Variant 
Public i As Integer 
Publip ss As AcadSelectionSet 

Sub points4d()
Dim objentity Аs Object
Randomize
s = 100 * Rnd()
Set ss = ThisDrawing.SelectionSets.Add(s)
filename = MsgBox("Выберите поверхность для фрезерования:", vbQuestion, "Что фрезеровать?") 
l1: ss.SelectOnScreen
If ss.Count > 1 Then 
	ss.Clear
	filename = MsgBox("Выберите ОДНУ поверхность!", vbExclamation, "Предупреждение") 
	GoTo l1 
End If 
For Each objentity In ss
	If objentity.EntityType = acPolymesh Then 
		setki.Add objentity 
	End If 
Next objentity 
UserForm1.Show
End Sub

REM ---------------------------- UserForm1 - 1 ----------------------------
Private Sub CommandButton1_Click() 
filename = TextBoxl.Text 
If filename = "" Then 
	Exit Sub
End If
CommandButton1.Enabled = False
UserForm1.MousePointer = fmMousePointerHourGlass 
DoEvents
Open filename For Output Shared As #1 
ProgressBar1.Visible = True 
For Each setka In setki
	For i = 1 To setka.МVertexCount
		If i Mod 2 <> 0 Then
			For j = 1 To setka.NVertexCount
				ind = (i - 1) * setka.NVertexCount + j 
				ind_x = ind * 3 - 3 
				ind_y = ind * 3 - 2 
				ind_z = ind * 3 - 1
				Print #1, "X"; Fix(1000 * setka.Coordinates(ind_x)); _			
						"Y"; Fix(1000 * setka.Coordinates(ind_y)); _
						"Z"; Fix(1000 * setka.Coordinates(ind_z))
				v% = i * j * 100 / setka.MVertexCotftvt / setka.NVertexCount 
				ProgressBar1.Value = Int(v%)
			Next j
		Else
			For j = setka.NVertexCount To 1 Step -1
				ind = (i - 1) * setka.NVertexCount + j 
				ind_x = ind * 3 — 3 
				ind_y = ind * 3 - 2 
				ind_z = ind * 3 - 1
				Print #1, "X"; Fix(1000 * setka.Coordinates(ind_x)); _
				"Y"; Fix(1000 * setka.Coordinates(ind_y)); _
				"Z"; Fix(1000 * setka.Coordinates(ind_z))
			Next j 
		End If 
	Next i 
Next setka 
Close #1
UserForm1.MousePointer = fmMousePointerDefault 
UserForm1.Hide
MsgBox "' " & filename & " ' - " & "файл сформирован!", vblnformation, "Готово”
End Sub

REM ---------------------------- ------------- ----------------------------
Private Sub UserForm1_Activate()
ProgressBar1.Visible = False
End Sub




