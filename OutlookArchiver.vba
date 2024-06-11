Sub RunPythonScript()
    Dim objShell As Object
    Set objShell = VBA.CreateObject("WScript.Shell")
    objShell.Run "C:\Users\Paulx\AppData\Local\Programs\Python\Python312\python.exe C:\Users\Paulx\OneDrive\Bureaublad\DocumentenNiels\Scripts\Python\OutlookArchiver.py"
    Set objShell = Nothing
End Sub
