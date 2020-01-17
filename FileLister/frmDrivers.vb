Public Class frmDrivers

    Dim sDPPath As String
    Dim bCDUtil As Boolean

    Private Sub frmDrivers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        UpdateListing()
        Dim TempDriverDir As String = System.Environment.GetEnvironmentVariable("TEMP") & "\MCC_UL\Drvs\"

        If Not Dir(TempDriverDir & "dpinst.exe", vbHidden Or vbSystem Or vbReadOnly) = "" Then
            sDPPath = TempDriverDir & "dpinst.exe"
            bCDUtil = False
        Else
            Dim Drvs As Collections.ObjectModel.ReadOnlyCollection(Of System.IO.DriveInfo) _
                = My.Computer.FileSystem.Drives
            For Each Drv As System.IO.DriveInfo In Drvs
                If Drv.DriveType = IO.DriveType.CDRom Then
                    Dim CDPath As String = Drv.Name
                    If Drv.IsReady Then
                        Dim TempFile As String = Dir(CDPath & "\Drivers\dpinst.exe")
                        If TempFile = "" Then
                            TempFile = Dir(CDPath & "\ICalUL\Drivers\dpinst.exe")
                            If Not TempFile = "" Then
                                sDPPath = CDPath & "\ICalUL\Drivers\dpinst.exe"
                                bCDUtil = True
                            End If
                        Else
                            sDPPath = CDPath & "\Drivers\dpinst.exe"
                            bCDUtil = True
                        End If
                    Else
                        bCDUtil = False
                    End If
                End If
            Next
            If Not bCDUtil Then
                MsgBox("You don't have access to the dpinst utility and " & _
                    "no CD drive containing the utility was detected.", vbOKOnly, _
                    "Utility Not Stored Locally")
                Me.lstDrivers.Enabled = False
            End If
        End If

    End Sub

    Private Sub cmdDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDone.Click

        Me.Hide()

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

        Dim Resp As MsgBoxResult
        Dim DBins, BinString, InfName As String

        If sDPPath = "" Then
            MsgBox("You don't have access to the dpinst utility " & _
                "and no CD drive detected to read from UL disk.", _
                vbOKOnly, "Utility Not Stored Locally")
            Exit Sub
        End If

        If bCDUtil Then
            If Dir(sDPPath) = "" Then
                MsgBox("You must have the UL CD in the drive to use this utility.", _
                    vbOKOnly, "Utility Not Stored Locally")
            End If
        End If

        DBins = "?"
        BinString = ""
        If chkIncludeBin.Checked Then
            DBins$ = " and the associated binaries?"
            BinString$ = " /d"
        End If
        InfName = Me.lstDrivers.SelectedItem.ToString
        Resp = MsgBox("Are you sure you want to uninstall " & _
           InfName$ & DBins$, vbYesNo, "Remove Driver?")
        If Resp = vbYes Then
            Dim InfPath As String = FileLister.sDvrList.Get(InfName)
            Dim rst As Integer = Shell(sDPPath & " /u " & InfPath$ & _
               "\" & InfName$ & BinString$, vbNormalNoFocus)
            FileLister.sDvrList.Remove(InfName)
            UpdateListing()
        End If

    End Sub

    Private Sub UpdateListing()

        lstDrivers.Items.Clear()
        For Each DvrName As String In FileLister.sDvrList
            lstDrivers.Items.Add(DvrName)
        Next
        If Not (lstDrivers.Items.Count = 0) Then
            Me.cmdRemove.Enabled = Not (lstDrivers.SelectedItem = "")
        End If

    End Sub

    Private Sub lstDrivers_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstDrivers.SelectedIndexChanged

        If Not (lstDrivers.SelectedItem = "") Then
            Me.cmdRemove.Enabled = True
        End If

    End Sub

End Class