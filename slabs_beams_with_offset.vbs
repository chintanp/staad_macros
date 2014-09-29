Sub Main()

'Try to make several STAAD files from the same file or make work easier
	'TODO: Add your code here

'Declare the variables

Dim objOpenSTAAD1 As Object
Dim objOpenSTAAD2 As Object
Dim strxAxis As String
Dim stryAxis As String
Dim strzAxis As String
Dim SelBeamsNo As Long
Dim SelBeams() As Long
Dim bResult As Boolean
Dim propertyNo As Long
Dim wbk As Variant
Dim txt As String
Dim Ydx As Double
Dim Zdx As Double
Dim Ydz As Double
Dim Zdz As Double
Dim EX1 As  Application
Dim Book1 As Workbook
Dim Sheet1 As Variant
Dim EX2 As  Application
Dim Book2 As Workbook
Dim Sheet2 As Variant
Dim plateCenterMomentsPos(0 To 2) As Double
Dim plateCenterMomentsNegLong(0 To 2) As Double
Dim plateCenterMomentsNegShort(0 To 2) As Double
Dim bIncludePath As Boolean
Dim strFileName As String
Dim pnResult As Integer
Dim i As Integer
Dim j As Integer
Dim S As Integer
Dim count As Integer
Dim OffZ As Double
Dim OffX As Double
Dim lSpecStartZ As Integer
Dim lSpecEndZ As Integer
Dim lSpecStartX As Integer
Dim lSpecEndX As Integer
Dim OffResult1 As Integer
Dim OffResult2 As Integer
Dim OffResult3 As Integer
Dim OffResult4 As Integer


count =1
S = 2



'Initialize the STAAD Appplication Object
Set objOpenSTAAD1 = GetObject( , "StaadPro.OpenSTAAD")

'Retrieve the Filename of the currently open staad file with full path
bIncludePath = True
objOpenSTAAD1.GetSTAADFile strFileName, bIncludePath


'Initialize the variables with direction names
strxAxis = "X"
stryAxis = "Y"
strzAxis = "Z"

'Initialize the Excel Object
 Set EX1 = CreateObject ( "Excel.Application")

 EX1.Visible = True

'Initialize the Workbook object
 Set Book1 = EX1.Workbooks.Open("D:\Dropbox\structural engg\Thesis\CP thesis\Final Presentation\Results for all Load Combinations\beam_data.xlsx")

  'All the functions are available with Application Book
  Set Sheet1 = Book1.ActiveSheet

'Initialize the Excel Object
 Set EX2 = CreateObject ( "Excel.Application")

 EX2.Visible = True

'Initialize the Workbook object
 Set Book2 = EX2.Workbooks.Open("D:\Dropbox\structural engg\Thesis\CP thesis\Final Presentation\Results for all Load Combinations\result_data46_with_beam_offset.xlsx")

 Set Sheet2 = Book2.ActiveSheet

For i = 1 To 9 Step 3

For j = i To i+2 Step 1

	'Initialize the STAAD Appplication Object
Set objOpenSTAAD1 = GetObject( , "StaadPro.OpenSTAAD")

'Retrieve the Filename of the currently open staad file with full path
bIncludePath = True
objOpenSTAAD1.GetSTAADFile strFileName, bIncludePath


	'objOpenSTAAD1.ShowApplication

   		Zdz = Sheet1.Cells(j,1)
   		Ydz = Sheet1.Cells(j, 2)
   		Zdx = Sheet1.Cells(i,3)
  		 Ydx = Sheet1.Cells(i,4)



'Select Members parallel to a certain direction
objOpenSTAAD1.View.SelectMembersParallelTo (strzAxis)

'Get the number of selected Beams
SelBeamsNo = objOpenSTAAD1.Geometry.GetNoOfSelectedBeams

'Reallocate
ReDim SelBeams(SelBeamsNo-1) As Long

'Get the selected beams
objOpenSTAAD1.Geometry.GetSelectedBeams (SelBeams, 1)

'Set Material to Concrete
objOpenSTAAD1.Property.SetMaterialID 2

'Create a new property
propertyNo = objOpenSTAAD1.Property.CreatePrismaticRectangleProperty (Ydz, Zdz)

'Assign the created property
bResult = objOpenSTAAD1.Property.AssignBeamProperty (SelBeams, propertyNo)

'Offset the Z Beams
OffZ = 0.135/2 - Ydz/2
lSpecStartZ = objOpenSTAAD1.Property.CreateMemberOffsetSpec(0, 0, 0, OffZ, 0)
lSpecEndZ = objOpenSTAAD1.Property.CreateMemberOffsetSpec(1, 0, 0, OffZ, 0)
OffResult1 = objOpenSTAAD1.Property.AssignMemberSpecToBeam(SelBeams, lSpecStartZ)
OffResult2 = objOpenSTAAD1.Property.AssignMemberSpecToBeam(SelBeams, lSpecEndZ)


'Update Structure
objOpenSTAAD1.UpdateStructure

'Select Members parallel to a certain direction
objOpenSTAAD1.View.SelectMembersParallelTo (strxAxis)

'Get the number of selected Beams
SelBeamsNo = objOpenSTAAD1.Geometry.GetNoOfSelectedBeams

'Reallocate
ReDim SelBeams(SelBeamsNo-1) As Long

'Get the selected beams
objOpenSTAAD1.Geometry.GetSelectedBeams (SelBeams, 1)

'Set Material to Concrete
objOpenSTAAD1.Property.SetMaterialID 2

'Create a new property
propertyNo = objOpenSTAAD1.Property.CreatePrismaticRectangleProperty (Ydx, Zdx)

'Assign the created property
bResult = objOpenSTAAD1.Property.AssignBeamProperty (SelBeams, propertyNo)

'Offset the X Beams
OffX = 0.135/2 - Ydx/2
lSpecStartX = objOpenSTAAD1.Property.CreateMemberOffsetSpec(0, 0, 0, OffX, 0)
lSpecEndX = objOpenSTAAD1.Property.CreateMemberOffsetSpec(1, 0, 0, OffX, 0)
OffResult3 = objOpenSTAAD1.Property.AssignMemberSpecToBeam(SelBeams, lSpecStartX)
OffResult4 = objOpenSTAAD1.Property.AssignMemberSpecToBeam(SelBeams, lSpecEndX)

'Update Structure
objOpenSTAAD1.UpdateStructure

' Set the analysis mode as slient
objOpenSTAAD1.SetSilentMode 1

'Analyze the structure
objOpenSTAAD1.Analyze

While 1
	pnResult = objOpenSTAAD1.Output.AreResultsAvailable
	If pnResult = 1 Then
		Exit While
	End If
Wend

Set objOpenSTAAD2 = CreateObject("OpenSTAAD.Output.1")

objOpenSTAAD2.SelectSTAADFile strFileName

'Writing the results in the Excel File

'Reading the results for specific plates
'Plate 2140 represents the Center Plate Load Case 7 represents No Load on Beams
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 7, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 7 represents No Load on Beams
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 7, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 7 represents No Load on Beams
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 7, plateCenterMomentsNegLong(0)



 Sheet2.Cells(S+j,1).Value = Ydz
 Sheet2.Cells(S+j,2).Value = Zdz
 Sheet2.Cells(S+j,3).Value = Ydx
 Sheet2.Cells(S+j,4).Value = Zdx
 Sheet2.Cells(S+j,7).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,14).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,21).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,28).Value = plateCenterMomentsNegShort(0) * 4.44

 'Plate 2140 represents the Center Plate Load Case 8 represents Load on All X Beams
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 8, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 8 represents Load on All X Beams
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 8, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 8 represents Load on All X Beams
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 8, plateCenterMomentsNegLong(0)

Sheet2.Cells(S+j,8).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,15).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,22).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,29).Value = plateCenterMomentsNegShort(0) * 4.44

  'Plate 2140 represents the Center Plate Load Case 9 represents Load on All Z Beams
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 9, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 9 represents Load on All Z Beams
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 9, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 9 represents Load on All Z Beams
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 9, plateCenterMomentsNegLong(0)


Sheet2.Cells(S+j,9).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,16).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,23).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,30).Value = plateCenterMomentsNegShort(0) * 4.44


  'Plate 2140 represents the Center Plate Load Case 10 represents Line Load on one X Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 10, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 10 represents Line Load on one X Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 10, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 10 represents  Line Load on one X Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 10, plateCenterMomentsNegLong(0)


Sheet2.Cells(S+j,10).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,17).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,24).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,31).Value = plateCenterMomentsNegShort(0) * 4.44


  'Plate 2140 represents the Center Plate Load Case 11 represents Line Load on one Z Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 11, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 11 represents Line Load on one Z Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 11, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 11 represents Line Load on one Z Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 11, plateCenterMomentsNegLong(0)


Sheet2.Cells(S+j,11).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,18).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,25).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,32).Value = plateCenterMomentsNegShort(0) * 4.44


 'Plate 2140 represents the Center Plate Load Case 12 represents Line Load on one X and one Z Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 12, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 12 represents Line Load on one X and one Z Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 12, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 12 represents Line Load on one X and one Z Beam.
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 12, plateCenterMomentsNegLong(0)


Sheet2.Cells(S+j,12).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,19).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,26).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,33).Value = plateCenterMomentsNegShort(0) * 4.44


 'Plate 2140 represents the Center Plate Load Case 13 represents Load on all Beams.
objOpenSTAAD2.GetAllPlateCenterMoments 2140, 13, plateCenterMomentsPos(0)

'Plate 2131 represents Mx- Load Case 13 represents Load on all Beams.
objOpenSTAAD2.GetAllPlateCenterMoments 2131, 13, plateCenterMomentsNegShort(0)

'Plate 1930 represents My- Load Case 13 represents Load on all Beams.
objOpenSTAAD2.GetAllPlateCenterMoments 1930, 13, plateCenterMomentsNegLong(0)


Sheet2.Cells(S+j,13).Value = plateCenterMomentsPos(0) * 4.44
 Sheet2.Cells(S+j,20).Value = plateCenterMomentsPos(1) * 4.44
 Sheet2.Cells(S+j,27).Value = plateCenterMomentsNegLong(1) * 4.44
 Sheet2.Cells(S+j,34).Value = plateCenterMomentsNegShort(0) * 4.44



 Book2.Save

'Closes the link to analysis so the post processing results are unavailable
objOpenSTAAD2.CloseAnalysisLink

Set objOpenSTAAD1 = Nothing
Set objOpenSTAAD2 = Nothing

'count = count + 1
Next
S = S + 3
'count = 1
Next


Set EX1 = Nothing
Set Book1 = Nothing
Set Sheet1 = Nothing
Set EX2 = Nothing
Set Book2 = Nothing
Set Sheet2 = Nothing


End Sub
