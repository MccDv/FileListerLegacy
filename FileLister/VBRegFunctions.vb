Imports Microsoft.Win32
Imports Microsoft.Win32.Registry
Imports Microsoft.VisualBasic

Module VBRegFunctions

    Public Function RegKeyExists(ByVal Hive As RegistryHive, ByVal NodeToTest As String) As Boolean

        Dim CurRegKey As RegistryKey
        Dim HiveKey As RegistryKey
        Dim KeyExists As Boolean = False

        HiveKey = GetHiveKey(Hive)
        CurRegKey = HiveKey.OpenSubKey(NodeToTest, False)
        If Not IsNothing(CurRegKey) Then
            KeyExists = True
            CurRegKey.Close()
        End If
        Return KeyExists

    End Function

    Public Function FindSubNodes(ByVal Hive As RegistryHive, ByRef _
        NodeToTest As String, ByRef ReturnedKeyNames As ArrayList) As Integer

        Dim iNumKeys As Integer, index As Integer
        Dim CurRegKey As RegistryKey
        Dim SubKeyNames() As String
        Dim HiveKey As RegistryKey

        HiveKey = GetHiveKey(Hive)
        CurRegKey = HiveKey.OpenSubKey(NodeToTest, False)
        If Not IsNothing(CurRegKey) Then
            SubKeyNames = CurRegKey.GetSubKeyNames()
            iNumKeys = SubKeyNames.Length
            If iNumKeys < 0 Then
                ReturnedKeyNames = Nothing
            Else
                For index = 0 To iNumKeys - 1
                    ReturnedKeyNames.Add(SubKeyNames(index))
                Next
            End If
            CurRegKey.Close()
        End If
        HiveKey.Close()
        Return iNumKeys

    End Function

    Public Function GetAllValues(ByVal Hive As RegistryHive, ByRef _
        NodeToTest As String, ByRef ReturnedValueNames As ArrayList) As Integer

        Dim iNumKeys As Integer, index As Integer
        Dim CurRegKey As RegistryKey
        Dim ValueNames() As String
        Dim HiveKey As RegistryKey

        HiveKey = GetHiveKey(Hive)
        CurRegKey = HiveKey.OpenSubKey(NodeToTest, False)
        If Not IsNothing(CurRegKey) Then
            ValueNames = CurRegKey.GetValueNames()
            iNumKeys = ValueNames.Length
            If iNumKeys < 0 Then
                ReturnedValueNames = Nothing
            Else
                For index = 0 To iNumKeys - 1
                    ReturnedValueNames.Add(ValueNames(index))
                Next
            End If
            CurRegKey.Close()
        End If
        HiveKey.Close()
        Return iNumKeys

    End Function

    Public Function GetKeyValue(ByVal Hive As RegistryHive, ByRef _
        RegNode As String, ByRef ValName As String) As Boolean

        Dim CurRegKey As RegistryKey
        Dim HiveKey As RegistryKey
        Dim ValFound As Boolean = False
        Dim ValueFound As String = ""

        HiveKey = GetHiveKey(Hive)
        CurRegKey = HiveKey.OpenSubKey(RegNode, False)
        If Not IsNothing(CurRegKey) Then
            ValueFound = CurRegKey.GetValue(ValName)
            If Not IsNothing(ValueFound) Then
                If Not ValueFound = "" Then
                    ValFound = True
                End If
            End If
            CurRegKey.Close()
        End If
        HiveKey.Close()
        ValName = ValueFound
        Return ValFound

    End Function

    Public Function SetKeyValue(ByVal Hive As RegistryHive, ByRef RegNode _
        As String, ByRef ValName As String, ByRef ValToSet As String) As Boolean

        Dim CurRegKey As RegistryKey
        Dim HiveKey As RegistryKey
        Dim ValSet As Boolean = False
        Dim ValueFound As String = ""

        HiveKey = GetHiveKey(Hive)
        Try
            CurRegKey = HiveKey.OpenSubKey(RegNode, True)
            If Not IsNothing(CurRegKey) Then
                CurRegKey.SetValue(ValName, ValToSet)
                ValSet = True
                CurRegKey.Close()
            End If
        Catch ex As UnauthorizedAccessException
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error Writing Registry")
        End Try
        HiveKey.Close()
        Return ValSet

    End Function

    Private Function GetHiveKey(ByVal Hive As RegistryHive) As RegistryKey

        Dim HiveKey As RegistryKey

        HiveKey = Nothing
        HiveKey = Switch( _
            Hive = RegistryHive.CurrentConfig, Registry.CurrentConfig, _
            Hive = RegistryHive.CurrentUser, Registry.CurrentUser, _
            Hive = RegistryHive.LocalMachine, Registry.LocalMachine, _
            Hive = RegistryHive.Users, Registry.Users, _
            Hive = RegistryHive.ClassesRoot, Registry.ClassesRoot)
        Return HiveKey

    End Function

End Module
