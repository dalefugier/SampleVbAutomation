Imports System
Imports System.Threading
Imports Rhino4
Imports RhinoScript4

Module MainModule

  Sub Main()

    'Const rhinoId = "Rhino5.Application";
    Const rhinoId As String = "Rhino5.Interface"
    'Const rhinoId As String = "Rhino5x64.Application"
    'Const rhinoId = "Rhino5x64.Interface";

    'Dim rhinoObj As Rhino5Application = Nothing
    Dim rhinoObj As Rhino5Interface = Nothing

    Try
      Dim rhinoType As Type = Type.GetTypeFromProgID(rhinoId)
      rhinoObj = Activator.CreateInstance(rhinoType)
    Catch
    End Try

    If (rhinoObj Is Nothing) Then
      Console.WriteLine("Failed to create: " & rhinoId)
      Return
    End If

    Const bailTime As Integer = 15 * 1000
    Dim waitTime As Integer = 0
    While (0 = rhinoObj.IsInitialized())
      Thread.Sleep(100)
      waitTime += 100
      If (waitTime > bailTime) Then
        Console.WriteLine("Rhino initialization failed")
        Return
      End If
    End While

    rhinoObj.Visible = 1

    Dim rsObj As RhinoScript = rhinoObj.GetScriptObject()
    If (rsObj Is Nothing) Then
      Console.WriteLine("Failed to create RhinoScript object")
      Return
    End If

    Dim strObj = rsObj.GetObject("Select a surface", 8)
    If (Not IsDBNull(strObj)) Then
      Console.WriteLine("ID: " & strObj)

      Dim arrMassProp = rsObj.SurfaceAreaMoments(strObj)
      If (Not IsDBNull(arrMassProp)) Then
        Dim str = rsObj.Pt2Str(arrMassProp(6), 4)
        Console.WriteLine("Area Moments of Inertia about the World Coordinate Axes: " & str)
      End If

    End If

  End Sub

End Module
