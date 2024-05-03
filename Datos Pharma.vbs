Dim fso: set fso = CreateObject("Scripting.FileSystemObject")
Dim CurrentDirectory
CurrentDirectory = fso.GetAbsolutePathName(".")
dim Directory
Directory = CurrentDirectory & "\PharmaData.dll"
Set ObjExcel = CreateObject("Excel.Application")
Objexcel.Visible = True
objexcel.width = 0
objexcel.height = 0
ObjExcel.Workbooks.Open Directory,0,true
Set ObjExcel = Nothing
Set ObjLibro = Nothing