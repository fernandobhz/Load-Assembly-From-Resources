Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.Reflection
Imports System.Security.Cryptography

Public Class EmbeddedAssembly

    Shared Dic As Dictionary(Of String, Assembly) = Nothing

    Private Shared isHandlerAdded As Boolean

    Public Shared Sub Load(DLLBuff As Byte())
        If Dic Is Nothing Then
            Dic = New Dictionary(Of String, Assembly)()
        End If


        If Not isHandlerAdded Then
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf EmbeddedAssembly.CurrentDomain_AssemblyResolve
            isHandlerAdded = True
        End If


        Dim asm = Assembly.Load(DLLBuff)
        Dic.Add(asm.FullName, asm)
    End Sub

    Public Shared Sub Load(DLLName As String) 'Eu que fiz

        If Not isHandlerAdded Then
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf EmbeddedAssembly.CurrentDomain_AssemblyResolve
            isHandlerAdded = True
        End If


        Dim embeddedResource As String = Assembly.GetExecutingAssembly().GetManifestResourceNames.FirstOrDefault(Function(x) x.Contains(DLLName))
        Load(embeddedResource, DLLName)
    End Sub

    Public Shared Function CurrentDomain_AssemblyResolve(sender As Object, args As ResolveEventArgs) As Assembly
        Return EmbeddedAssembly.Get(args.Name)
    End Function

    ''' <summary>
    ''' Load Assembly, DLL from Embedded Resources into memory.
    ''' </summary>
    ''' <param name="embeddedResource">Embedded Resource string. Example: WindowsFormsApplication1.SomeTools.dll</param>
    ''' <param name="fileName">File Name. Example: SomeTools.dll</param>
    Public Shared Sub Load(embeddedResource As String, fileName As String)
        If Dic Is Nothing Then
            Dic = New Dictionary(Of String, Assembly)()
        End If

        Dim ba As Byte() = Nothing
        Dim asm As Assembly = Nothing
        Dim curAsm As Assembly = Assembly.GetExecutingAssembly()

        Using stm As Stream = curAsm.GetManifestResourceStream(embeddedResource)
            ' Either the file is not existed or it is not mark as embedded resource
            If stm Is Nothing Then
                Throw New Exception(embeddedResource & Convert.ToString(" is not found in Embedded Resources."))
            End If

            ' Get byte[] from the file from embedded resource
            ba = New Byte(CInt(stm.Length) - 1) {}
            stm.Read(ba, 0, CInt(stm.Length))
            Try
                asm = Assembly.Load(ba)

                ' Add the assembly/dll into dictionary
                Dic.Add(asm.FullName, asm)
                Return
                ' Purposely do nothing
                ' Unmanaged dll or assembly cannot be loaded directly from byte[]
                ' Let the process fall through for next part
            Catch
            End Try
        End Using

        Dim fileOk As Boolean = False
        Dim tempFile As String = ""

        Using sha1 As New SHA1CryptoServiceProvider()
            ' Get the hash value from embedded DLL/assembly
            Dim fileHash As String = BitConverter.ToString(sha1.ComputeHash(ba)).Replace("-", String.Empty)

            ' Define the temporary storage location of the DLL/assembly
            tempFile = Path.GetTempPath() & fileName

            ' Determines whether the DLL/assembly is existed or not
            If File.Exists(tempFile) Then
                ' Get the hash value of the existed file
                Dim bb As Byte() = File.ReadAllBytes(tempFile)
                Dim fileHash2 As String = BitConverter.ToString(sha1.ComputeHash(bb)).Replace("-", String.Empty)

                ' Compare the existed DLL/assembly with the Embedded DLL/assembly
                If fileHash = fileHash2 Then
                    ' Same file
                    fileOk = True
                Else
                    ' Not same
                    fileOk = False
                End If
            Else
                ' The DLL/assembly is not existed yet
                fileOk = False
            End If
        End Using

        ' Create the file on disk
        If Not fileOk Then
            System.IO.File.WriteAllBytes(tempFile, ba)
        End If

        ' Load it into memory
        asm = Assembly.LoadFile(tempFile)

        ' Add the loaded DLL/assembly into dictionary
        Dic.Add(asm.FullName, asm)
    End Sub

    ''' <summary>
    ''' Retrieve specific loaded DLL/assembly from memory
    ''' </summary>
    ''' <param name="assemblyFullName"></param>
    ''' <returns></returns>
    Public Shared Function [Get](assemblyFullName As String) As Assembly
        If Dic Is Nothing OrElse Dic.Count = 0 Then
            Return Nothing
        End If

        If Dic.ContainsKey(assemblyFullName) Then
            Return Dic(assemblyFullName)
        End If

        Return Nothing

        ' Don't throw Exception if the dictionary does not contain the requested assembly.
        ' This is because the event of AssemblyResolve will be raised for every
        ' Embedded Resources (such as pictures) of the projects.
        ' Those resources wil not be loaded by this class and will not exist in dictionary.
    End Function
End Class
