Sub Makro()

    Const swDocPART = 1         
    Const swDocASSEMBLY = 2
    Const swDocDRAWING = 3
 
    Dim swApp As Object 
    Dim Part As Object 
    Dim face As Object 
    'Dim massProps As Variant 
    Dim R1, R2, L, a, b, f As Double
    Dim i As Integer
    R1min = 10: R1max = 50: R1step = 5
    R2min = 10: R2max = 50: R2step = 5
    Lmin = 20: Lmax = 60: Lstep = 10
    amin = 10: amax = 50: astep = 5
    bmin = 15: bmax = 50: bstep = 5
    fmin = 1: fmax = 10: fstep = 2
    MyPath = CurDir
    MyPath = "C:\Protector"

    Set swApp = CreateObject("SldWorks.Application")
    Set Part = swApp.OpenDoc(MyPath + "\ModelVBA.SLDPRT", swDocPART)
    If Part Is Nothing Then
       Exit Sub
    Else
       Set Part = swApp.ActivateDoc("ModelVBA.SLDPRT")
    End If
    i = 3 
    For R1 = R1min To R1max Step R1step
    For R2 = R2min To R2max Step R2step
    For L = Lmin To Lmax Step Lstep
    For a = amin To amax Step astep
    For b = bmin To bmax Step bstep
    For f = fmin To fmax Step fstep
    Part.Parameter("D1@Filet1").SystemValue = R1 / 1000
    Part.Parameter("D1@Filet2").SystemValue = R2 / 1000
    Part.Parameter("D1@Extrude2").SystemValue = L / 1000
    Part.Parameter("D1@c_sketch").SystemValue = (a * 3.14) / 180
    Part.Parameter("D3@schemfer").SystemValue = b / 1000
    Part.Parameter("D1@w_sketch").SystemValue = 0.056 - (f / 1000)
    Part.EditRebuild 
    Cells(i, 1).Value = R1: Cells(2, 1).Value = R1
    Cells(i, 2).Value = R2: Cells(2, 2).Value = R2
    Cells(i, 3).Value = L: Cells(2, 3).Value = L
    Cells(i, 4).Value = a: Cells(2, 4).Value = a
    Cells(i, 5).Value = b: Cells(2, 5).Value = b
    Cells(i, 6).Value = f: Cells(2, 6).Value = f
 
    'massProps = Part.GetMassProperties
    'Cells(i, 7).Value = 2 * massProps(3)

    Set face = Part.GetEntityByName(1, 2)
    Cells(i, 8).Value = face.GetArea
   Set face = Part.GetEntityByName(2, 2)
    Cells(i, 9).Value = 2 * face.GetArea
    i = i + 1 
    Next f 
    Next b
    Next a
    Next L
    Next R2
    Next R1
End Sub

