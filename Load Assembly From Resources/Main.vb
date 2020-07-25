Imports System.Reflection

Public Module Main

    Sub Main()
        EmbeddedAssembly.Load(My.Resources.DirectShowLib_2005)
        Form1.ShowDialog()
    End Sub


End Module
