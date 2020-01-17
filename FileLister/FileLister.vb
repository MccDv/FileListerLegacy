Imports Microsoft.Win32
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Diagnostics
Imports System.Collections.SortedList
Imports System.Collections.Specialized

Public Class FileLister

    Dim sCurVolume As String = " "
    Dim sAndroidTempPath As String = ""
    Dim sBaseSW32 As String = "SOFTWARE\"
    Dim sBaseSW64 As String = "SOFTWARE\Wow6432Node\"
    Dim sWinUInst As String = "Microsoft\Windows\CurrentVersion\Uninstall"
    Dim sDocPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
    Dim sPublicDocPath As String = Environment.GetEnvironmentVariable("PUBLIC") & "\Documents\"
    'Dim sAllUserPath As String = Environment.GetEnvironmentVariable("ALLUSERSPROFILE") & "\Start Menu\"
    Dim sSystemPath As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
    Dim sWinPath As String = Environment.GetEnvironmentVariable("SystemRoot")
    Dim sVolumeRegKey As String = "Software\Measurement Computing\Universal Test Suite"
    Dim sVolumeReg32Key As String = "Software\Wow6432Node\Measurement Computing\Universal Test Suite"
    Dim cProgGroupNodes As New NameValueCollection
    Dim cMCCProgNodes As New NameValueCollection
    Dim cOddRegNodes As New NameValueCollection
    Dim cMCCProgGUIDs As New NameValueCollection
    Dim sbContents As New StringBuilder()
    Dim osVer As Version = Environment.OSVersion.Version
    Dim osPltfm As PlatformID = Environment.OSVersion.Platform
    Dim sOSBits As String
    Public sDvrList As New NameValueCollection


    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click

        Dim CurProgGroup As String, CurRegNode As String
        Dim ConfirmedRegNode As String, ValueFound As String = ""
        Dim NodesFound As Integer, RegSection As String
        Dim ReturnedKeyNames As New ArrayList
        Dim PlatformsFound As New ArrayList
        Dim CurProgram As String, TargetFile As String
        Dim OSVersion As String = "", OSShort As String = ""
        Dim ShortName As String = "", RegBase As String
        Dim Response As MsgBoxResult = MsgBoxResult.Cancel
        Dim BaseName As String

        If Me.txtOutput.TextLength > 50 Then
            Response = MsgBox("Listing has not been saved. " & _
                "Clear listing?", MsgBoxStyle.OkCancel, "Replace Listing?")
            If Response = MsgBoxResult.Cancel Then Exit Sub
        Else
            If sCurVolume = " " Then
                Response = MsgBox("CD Image has not been saved. " & _
                    "Imaging the CD is recommended before saving any listings.", _
                    MsgBoxStyle.Information, "Replace Listing?")
                sCurVolume = ""
            End If
        End If
        lstKeys.Items.Clear()
        lstValues.Items.Clear()
        cProgGroupNodes.Clear()
        ConfirmedRegNode = ""
        Dim Prefix As String = "\Env "
        Dim ModGroup As String
        For Each ProgGroup As String In My.Settings.MCCProgGroups
            CurProgGroup = ProgGroup
            BaseName = ProgGroup
            ModGroup = ProgGroup
            ConfirmedRegNode = sBaseSW32 & ProgGroup
            RegBase = ConfirmedRegNode
            If ProgGroup.Contains("|") Then
                Dim SplitGroup As Array = ProgGroup.Split("|")
                ProgGroup = SplitGroup(1)
                BaseName = SplitGroup(0)
                ModGroup = ProgGroup
                If BaseName.Contains("DASYLab") Then
                    If BaseName.EndsWith("DASYLab") Then
                        'ModGroup = ProgGroup.Remove(ProgGroup.IndexOf("\"))
                        If ProgGroup.Contains("National Instruments") Then
                            BaseName = BaseName.Insert(0, "NI ")
                        End If
                        'Else
                        '   ModGroup = ProgGroup
                    End If
                    RegBase = sBaseSW32 & ModGroup
                End If
                ConfirmedRegNode = sBaseSW32 & ProgGroup
            End If
            RegSection = " (64)"
            If sOSBits = "32-bit" Then RegSection = " (32)"
            For i As Integer = 0 To 1
                If RegKeyExists(RegistryHive.LocalMachine, ConfirmedRegNode) Then
                    cProgGroupNodes.Add(BaseName & RegSection, RegBase)
                    lstKeys.Items.Add(BaseName & RegSection)
                End If
                ConfirmedRegNode = sBaseSW64 & ProgGroup
                RegBase = sBaseSW64 & ModGroup
                RegSection = " (32)"
            Next
        Next
        txtFileName.Clear()
        sbContents.Length = 0
        sbContents.AppendLine("FileLister for .Net version: " & Application.ProductVersion)
        Dim SPack As String = Environment.OSVersion.ServicePack
        sbContents.AppendLine(GetPCName(ShortName) & "      " & _
            GetCurOS(OSVersion, OSShort) & SPack)
        sbContents.AppendLine(Environment.GetEnvironmentVariable("PATH"))
        Dim WinInst As String = sSystemPath & "\" & "msiexec.exe"
        If IO.File.Exists(WinInst) Then
            Dim fvi As FileVersionInfo = FileVersionInfo.GetVersionInfo(WinInst)
            sbContents.AppendLine("Windows Installer version " & fvi.FileVersion)
        End If
        sbContents.AppendLine()
        Dim Insertion As Integer = sbContents.Length
        Dim PRFound As Boolean = False
        Dim ProdPrefix As String = ""
        Dim RegSpec As String, PathString As String
        For Each KeyID As String In cProgGroupNodes.AllKeys
            CurRegNode = cProgGroupNodes.Item(KeyID)
            ReturnedKeyNames.Clear()
            NodesFound = FindSubNodes(RegistryHive.LocalMachine, CurRegNode, ReturnedKeyNames)
            For Each rkn As String In ReturnedKeyNames
                CurProgram = rkn
                RegSpec = CurRegNode & "\" & CurProgram
                If SearchRevNames(RegSpec, ValueFound) Then
                    Dim ValName As String = "ProductName"
                    If GetKeyValue(RegistryHive.LocalMachine, RegSpec, ValName) Then
                        If Not ValName = "" Then
                            CurProgram = ValName
                            PathString = ""
                            'Dim PathFound As Boolean = SearchPathNames(RegSpec, PathString)
                            'If PathFound Then 
                            'cProgGroupNodes.Set(KeyID, RegSpec)
                        End If
                    End If
                    If CurRegNode.EndsWith("DASYLab Drivers") Then ProdPrefix = "DASYLab Drivers "
                    sbContents.AppendLine("  " & ProdPrefix & CurProgram & "  " & ValueFound)
                    PRFound = True
                    ProdPrefix = ""
                End If
            Next
        Next

        Dim AndroidVersion As String = ""
        AndroidVersion = GetAndroidVersion()

        If Not AndroidVersion = "" Then sbContents.AppendLine("  UL for Android  " & AndroidVersion)
        If PRFound Then sbContents.Insert(Insertion, "Installed MCC Applications" & vbCrLf)
        Insertion = sbContents.Length
        Dim DEFound As Boolean = GetDevelopmentEnvirons()
        If DEFound Then sbContents.Insert(Insertion, "Development Environments" & vbCrLf)
        Insertion = sbContents.Length
        Dim PFFound As Boolean = SearchForPlatforms(PlatformsFound)
        If PFFound Then sbContents.Insert(Insertion, "Platforms" & vbCrLf)
        Insertion = sbContents.Length
        Dim DepFound As Boolean = GetDependencies()
        If DepFound Then sbContents.Insert(Insertion, "Dependencies" & vbCrLf)
        Insertion = sbContents.Length
        Dim OEFound As Boolean = GetOtherEntries()
        If OEFound Then sbContents.Insert(Insertion, "Other MCC Registry Entries" & vbCrLf)
        sbContents.AppendLine()
        If Not (sCurVolume.Contains(OSShort)) Then sCurVolume _
            = sCurVolume & " " & OSShort & " " & ShortName
        TargetFile = sDocPath & Prefix & sCurVolume & ".txt"

        UpdateTextList(TargetFile, True)

    End Sub

    Private Sub cmdImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImage.Click

        Dim DirList As New ArrayList
        Dim TargetFile As String
        Dim Filter As String = "*.*"
        Dim vl As String = ""

        If Me.txtOutput.TextLength > 50 Then
            Dim Response As MsgBoxResult = MsgBox("Listing has not been saved. " & _
                "Clear listing?", MsgBoxStyle.OkCancel, "Replace Listing?")
            If Response = MsgBoxResult.Cancel Then Exit Sub
        End If
        sbContents.Length = 0
        Dim Result As DialogResult = FolderBrowserDialog1.ShowDialog()
        If Result = Windows.Forms.DialogResult.OK Then
            Dim SelFolder As String = FolderBrowserDialog1.SelectedPath
            Dim di As DriveInfo = My.Computer.FileSystem.GetDriveInfo(SelFolder)
            If Not di.DriveType = DriveType.CDRom Then
                Dim Warn As DialogResult = MsgBox(SelFolder & _
                    " is not the CD ROM. Get image anyway?", _
                    MsgBoxStyle.YesNo, "CD ROM Expected")
                If Warn = Windows.Forms.DialogResult.No Then Exit Sub
                vl = SelFolder.Substring(SelFolder.LastIndexOf("\") + 1)
            Else
                vl = di.VolumeLabel
            End If
            sCurVolume = vl
            Dim CurDate As New Date
            CurDate = Now
            sbContents.AppendLine("Date of listing: " & CurDate.ToString("dddd, MMMM dd, yyyy"))
            If Not SetKeyValue(RegistryHive.CurrentUser, sVolumeRegKey, "CurVolume", vl) Then _
                SetKeyValue(RegistryHive.CurrentUser, sVolumeReg32Key, "CurVolume", vl)
            sbContents.AppendLine()
            CreateDirListing(SelFolder, Filter, True)
            GetDirectories(SelFolder, DirList)
            For Each Dir As String In DirList
                CreateDirListing(Dir, Filter)
            Next
        End If
        TargetFile = sDocPath & "\Img " & vl & ".txt"

        UpdateTextList(TargetFile, True)

    End Sub

    Private Sub cmdListFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles cmdListFiles.Click, cmdListProps.Click

        Dim SelString As String
        Dim DirList As New ArrayList
        Dim TargetFile As String
        Dim FilterString As String = "*.*"
        Dim TempString() As String, Trailer As Integer
        Dim ListProps As Boolean = sender.Equals(cmdListProps)
        Dim Response As MsgBoxResult = MsgBoxResult.Cancel
        Dim ShareGroup As Collections.Specialized.StringCollection

        Dim Prefix As String = "\dir "

        ShareGroup = My.Settings.SharedFiles
        ShareGroup = GetSelectedPackage("Shares")
        TargetFile = txtFileName.Text
        If ListProps Then
            Prefix = "\props "
            If TargetFile.Contains("\dir ") Then
                Response = MsgBox("Directory listing has not been " & _
                    "saved. Clear listing?", MsgBoxStyle.OkCancel, _
                    "Replace Listing?")
                If Response = MsgBoxResult.Cancel Then Exit Sub
            End If
            'txtFileName.Clear()
        Else
            If Not (TargetFile.Contains("\dir ") Or TargetFile = "") Then
                Response = MsgBox("File listing is usually stored separately. " & _
                    "Append file listing to existing list?", MsgBoxStyle.OkCancel, _
                    "Append Listing?")
                If Response = MsgBoxResult.Cancel Then Exit Sub
            End If
            'txtFileName.Clear()
        End If
        sbContents.Length = 0
        Dim ProgParam As String
        Dim PathParam As String = lstValNames.SelectedItem.ToString
        If lstValues.Items.Count > 0 Then
            ProgParam = lstValues.SelectedItem.ToString
        Else
            Dim TempProg As String = Me.lstKeys.SelectedItem.ToString
            ProgParam = TempProg.Remove(TempProg.IndexOf(" ("))
        End If

        Dim ShareFilter As String

        If ListProps Then
            SelString = sSystemPath
            ShareFilter = ""
            For Each SharedFile As String In ShareGroup
                TempString = SharedFile.Split("|")
                If ProgParam.Contains(TempString(0)) Then
                    ShareFilter = TempString(0)
                    If TempString.Length > 0 Then ShareFilter = TempString(1)
                    CreatePropsListing(SelString, ShareFilter, ProgParam)
                    SelString = "SharedFiles"
                End If
            Next
            Dim ShareList As String = sbContents.ToString
            If ShareList.EndsWith(sSystemPath & vbCrLf & vbCrLf) Then
                sbContents.AppendLine("   No DLLs found in " & sSystemPath)
                sbContents.AppendLine()
            Else
                If Not (ShareList = "") Then sbContents.AppendLine()
            End If
        End If
        SelString = ParseFileFilter(PathParam, FilterString)
        TempString = SelString.Split("\")
        Dim SplitSize As Integer = TempString.Length
        Trailer = 1
        If SelString.Length - SelString.LastIndexOf("\") < 2 Then Trailer = 2
        TargetFile = sDocPath & Prefix & _
            TempString(SplitSize - Trailer) & " " & sCurVolume & ".txt"

        If ListProps Then
            CreatePropsListing(SelString, FilterString)
            For Each SubDir As String In My.Settings.ProgramSubDirs
                TempString = SubDir.Split("|")
                If ProgParam.Contains(TempString(0)) Then
                    Dim ProgSubDir As String = SelString & TempString(1)
                    CreatePropsListing(ProgSubDir, FilterString, ProgParam, False)
                End If
            Next
        Else
            sbContents.AppendLine("List of directory " & SelString & vbCrLf)
            CreateDirListing(SelString, FilterString, True)
            GetDirectories(SelString, DirList)
            For Each Dir As String In DirList
                CreateDirListing(Dir, FilterString, False)
            Next
        End If

        UpdateTextList(TargetFile, False)

    End Sub

    Private Sub cmdStartList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStartList.Click

        Dim DirList As New ArrayList
        Dim ReturnedRegNodes As New ArrayList
        Dim FoundMenuLoc As Boolean
        Dim TargetFile, RegNode, ValueFound As String
        Dim FilterString As String = "*.*"
        Dim PathParam, UserNames, MenuFilter As String

        sbContents.Length = 0
        RegNode = "SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\User Shell Folders\"
        ValueFound = "Common Programs"
        FoundMenuLoc = GetKeyValue(RegistryHive.LocalMachine, RegNode, ValueFound)
        TargetFile = sDocPath & _
            "\Menu " & sCurVolume & ".txt"

        'all users
        If FoundMenuLoc Then
            PathParam = ValueFound 'sAllUserPath
            CreateDirListing(PathParam, "MCCMenu")
            GetDirectories(PathParam, DirList)
            For Each Dir As String In DirList
                CreateDirListing(Dir, "MCCMenu")
            Next
        End If

        'get all possible user names
        UserNames = ""
        RegNode = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\"
        Dim NodesFound As Integer = FindSubNodes(RegistryHive.LocalMachine, RegNode, ReturnedRegNodes)
        For Each Node As String In ReturnedRegNodes
            ValueFound = "ProfileImagePath"
            FoundMenuLoc = GetKeyValue(RegistryHive.LocalMachine, RegNode & Node, ValueFound)
            If FoundMenuLoc Then
                UserNames = UserNames & ValueFound.Substring(ValueFound.LastIndexOf("\") + 1) & "|"
            End If
        Next
        UserNames = UserNames.Remove(UserNames.LastIndexOf("|"), 1)

        'current user
        PathParam = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
        DirList.Clear()
        MenuFilter = "MCCMenu|" & UserNames
        CreateDirListing(PathParam, MenuFilter)
        GetDirectories(PathParam, DirList)
        For Each Dir As String In DirList
            CreateDirListing(Dir, MenuFilter)
        Next

        UpdateTextList(TargetFile, False)

    End Sub

    Private Sub cmdListGac_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdListGac.Click

        Dim DirList As New ArrayList
        Dim TargetFile As String
        Dim FilterString As String = "*.*"
        Dim PathParam As String

        sbContents.Length = 0
        TargetFile = sDocPath & _
            "\GAC " & sCurVolume & ".txt"

        PathParam = sWinPath & "\Assembly"
        CreatePropsListing(PathParam, "MCCGac")
        GetDirectories(PathParam, DirList)
        For Each Dir As String In DirList
            CreatePropsListing(Dir, "MCCGac")
        Next

        UpdateTextList(TargetFile, False)

    End Sub

    Private Sub cmdDriverList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDriverList.Click

        FindDrivers()

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim TargetFile As String
        Dim TextExists As Boolean

        TargetFile = txtFileName.Text
        Using DirFile As New StreamWriter(TargetFile)
            DirFile.Write(txtOutput.Text)
        End Using
        txtFileName.Clear()
        lstValNames.Items.Clear()
        cmdListFiles.Enabled = False
        cmdListProps.Enabled = False
        sbContents.Length = 0
        txtOutput.Clear()
        TextExists = (txtOutput.TextLength > 0)
        Me.cmdSave.Enabled = TextExists
        Me.cmdClear.Enabled = TextExists

    End Sub

    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click

        Dim TextExists As Boolean

        sbContents.Length = 0
        txtOutput.Clear()
        TextExists = (txtOutput.TextLength > 0)
        cmdSave.Enabled = TextExists
        cmdClear.Enabled = TextExists
        txtFileName.Clear()

    End Sub

    Private Sub lstKeys_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstKeys.SelectedIndexChanged

        Dim CurKeyName As String, CurRegNode As String
        Dim CurProgram As String = ""
        Dim NodesFound As Integer, ItemFound As Boolean
        Dim CodeFound As String = ""
        Dim ValName As String, ListName As String
        Dim Include As Boolean
        Dim ReturnedKeyNames As New ArrayList

        cmdListFiles.Enabled = False
        cmdListProps.Enabled = False
        lstValues.Items.Clear()
        lstValNames.Items.Clear()
        cMCCProgNodes.Clear()
        If lstKeys.SelectedIndex < 0 Then Exit Sub
        cmdDriverList.Enabled = True
        cmdListGac.Enabled = True
        cmdStartList.Enabled = True
        CurKeyName = lstKeys.SelectedItem.ToString
        CurRegNode = cProgGroupNodes.Item(CurKeyName)
        NodesFound = FindSubNodes(RegistryHive.LocalMachine, CurRegNode, ReturnedKeyNames)
        For Each rkn As String In ReturnedKeyNames
            CurProgram = rkn
            ListName = CurProgram
            ValName = "ProductType"
            If GetKeyValue(RegistryHive.LocalMachine, CurRegNode & _
                "\" & CurProgram, ValName) Then
                If Not ValName = "" Then
                    ListName = ValName
                End If
            Else
                ValName = "ProductName"
                If GetKeyValue(RegistryHive.LocalMachine, CurRegNode & _
                    "\" & CurProgram, ValName) Then
                    If Not ValName = "" Then
                        ListName = ValName
                    End If
                End If
            End If
            Include = True
            For Each FilterParam As String In My.Settings.ProgFilter
                Dim Filter As String, Param As String
                Dim Splitter As Integer = FilterParam.IndexOf("|")
                Filter = FilterParam.Substring(0, Splitter - 1)
                Param = FilterParam.Substring(Splitter + 1)
                If CurRegNode.Contains(Filter) Then
                    Include = False
                    If CurRegNode.Contains(Param) Then
                        Include = True
                        Exit For
                    End If
                Else
                    Include = True
                End If
            Next
            If Include Then
                If CurProgram.Contains("Omni CD") Then
                    'lstValues.Items.Add("Data Translation Open Layers (OEM)")
                    'lstValues.Items.Add("Data Translation Open Layers")
                    lstValues.Items.Add("QuickDAQ")
                    cMCCProgNodes.Add("QuickDAQ", CurRegNode.Replace("Omni CD", "QuickDAQ"))
                End If
                lstValues.Items.Add(ListName)
                cMCCProgNodes.Add(ListName, CurRegNode & "\" & CurProgram)
            End If
        Next
        If CurProgram = "" Then
            ValName = "ProductType"
            If GetKeyValue(RegistryHive.LocalMachine, CurRegNode, ValName) Then
                If Not ValName = "" Then
                    CurProgram = ValName
                    lstValues.Items.Add(CurProgram)
                    cMCCProgNodes.Add(CurProgram, CurRegNode)
                End If
            Else
                ValName = "ProductName"
                If GetKeyValue(RegistryHive.LocalMachine, CurRegNode, ValName) Then
                    If Not ValName = "" Then
                        CurProgram = ValName
                        lstValues.Items.Add(CurProgram)
                        cMCCProgNodes.Add(CurProgram, CurRegNode)
                    End If
                End If
            End If
        End If
        If NodesFound = 0 Then
            ItemFound = SearchRevNames(CurRegNode, CodeFound)
            If ItemFound Then Me.lstValNames.Items.Add("Program revision " & CodeFound)
            ItemFound = SearchPathNames(CurRegNode, CodeFound)
            If ItemFound Then Me.lstValNames.Items.Add(CodeFound)
            ItemFound = SearchForProductCode(CurRegNode, CurKeyName, CodeFound)
            If ItemFound Then Me.lstValNames.Items.Add("Package Version GUID:  " & CodeFound)
        End If

    End Sub

    Private Sub lstValues_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstValues.SelectedIndexChanged

        Dim CurProgName As String, CurProgGroup As String
        Dim NumRegNodes As Integer, NumSubDirs As Integer
        Dim CurRegNode As String, ItemFound As Boolean
        Dim CodeFound As String = "", CurPath As String
        Dim PathParam As String
        Dim PathString As String, FilterString As String = "*.*"
        Dim TextExists As Boolean
        Dim RegNodes As New ArrayList
        Dim SubDirs As New ArrayList
        Dim InstallerDirs As New ArrayList
        Dim PublicDirs As New ArrayList, SampleDirsFound As Boolean

        If lstValues.SelectedIndex < 0 Then Exit Sub
        cmdListFiles.Enabled = False
        cmdListProps.Enabled = False
        'txtFileName.Text = ""
        lstValNames.Items.Clear()
        sbContents.Length = 0
        'txtOutput.Clear()
        cMCCProgGUIDs.Clear()
        TextExists = (txtOutput.TextLength > 0)
        Me.cmdSave.Enabled = TextExists
        Me.cmdClear.Enabled = TextExists

        CurProgName = lstValues.SelectedItem.ToString
        CurRegNode = cMCCProgNodes.Item(lstValues.SelectedItem.ToString)
        NumRegNodes = CheckForCustomRegNodes(CurProgName, RegNodes)
        NumSubDirs = CheckForCustomSubDirs(CurProgName, SubDirs)

        If NumRegNodes > 0 Then
            For Each rn As String In RegNodes
                For Each sd As String In SubDirs
                    PathParam = rn & sd
                    PathString = ParseFileFilter(PathParam, FilterString)
                    CurPath = PathString
                    If Directory.Exists(CurPath) Then _
                        Me.lstValNames.Items.Add(PathParam)
                Next
            Next
        ElseIf NumSubDirs > 0 Then
            For Each sd As String In SubDirs
                CurPath = sd
                If Directory.Exists(CurPath) Then _
                    Me.lstValNames.Items.Add(CurPath)
            Next
        End If

        CurProgGroup = lstKeys.SelectedItem.ToString
        'CurRegNode = cProgGroupNodes(CurProgGroup) & "\" & CurProgName
        ItemFound = SearchPathNames(CurRegNode, CodeFound)
        If Not ItemFound Then
            Dim rkn As New ArrayList
            Dim snf As Integer = FindSubNodes(RegistryHive.LocalMachine, CurRegNode, rkn)
            For Each returnedKey As String In rkn
                ItemFound = SearchPathNames(CurRegNode & "\" & returnedKey, CodeFound)
                If ItemFound Then Exit For
            Next
        End If
        If ItemFound Then Me.lstValNames.Items.Add(CodeFound)

        CurPath = ""
        If SearchInstallerPaths(CurProgName, InstallerDirs) Then
            For Each CurPath In InstallerDirs
                Dim DontList As Boolean = CurPath.Contains("Start Menu")
                DontList = DontList Or CurPath.Contains("Examples")
                If Not DontList Then lstValNames.Items.Add(CurPath)
            Next
        End If
        Dim KeyNames() As String = cOddRegNodes.AllKeys
        For Each KeyName As String In KeyNames
            If CurProgName = KeyName Then
                Dim OddRegNode As String = cOddRegNodes.Get(CurProgName)
                If OddRegNode.Contains(":\") Then
                    ItemFound = True
                    lstValNames.Items.Add(OddRegNode)
                Else
                    ItemFound = SearchPathNames(OddRegNode, CodeFound)
                    If ItemFound Then lstValNames.Items.Add(CodeFound)
                End If
            End If
        Next

        Dim IncludePublicFolders As Boolean
        IncludePublicFolders = (CurProgName = "Instacal & Universal Library") _
            Or (CurProgName = "InstaCal")
        If IncludePublicFolders Then _
            SampleDirsFound = SearchUserDirs(PublicDirs)
        If SampleDirsFound Then
            For Each pd As String In PublicDirs
                lstValNames.Items.Add(pd)
            Next
        End If
        ItemFound = SearchRevNames(CurRegNode, CodeFound)
        If ItemFound Then Me.lstValNames.Items.Add("Program revision " & CodeFound)
        ItemFound = SearchForProductCode(CurRegNode, CurProgName, CodeFound)
        If ItemFound Then
            Me.lstValNames.Items.Add("Package Version GUID:  " & CodeFound)
            cMCCProgGUIDs.Add(CurProgName, CodeFound)
        End If

    End Sub

    Private Sub lstValNames_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstValNames.MouseClick

        If Not lstValNames.SelectedIndex < 0 Then
            Clipboard.SetText(lstValNames.SelectedItem.ToString)
        End If

    End Sub

    Private Sub lstValNames_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstValNames.SelectedIndexChanged

        Dim PathString As String
        Dim PathParam As String, FilterString As String = "*.*"

        If lstValNames.SelectedIndex < 0 Then Exit Sub
        PathParam = lstValNames.SelectedItem.ToString
        PathString = ParseFileFilter(PathParam, FilterString)
        If PathString.Contains("\DAQ") Or PathString.Contains("\RedLab") Then
            If File.Exists(PathString & "Android\ul.jar") Then
                cmdUnZip.Enabled = True
                If sAndroidTempPath = "" Then
                    cmdUnZip.Text = "Unzip Files"
                Else
                    cmdUnZip.Text = "Del Archive"
                End If
            End If
        Else
            Me.cmdUnZip.Enabled = False
        End If
        Dim EnableCmds As Boolean = IO.Directory.Exists(PathString)
        cmdListFiles.Enabled = EnableCmds
        cmdListProps.Enabled = EnableCmds

    End Sub

    Private Function GetRegNodes(ByRef NodeToTest As String, ByRef _
        ReturnedRegNodes As ArrayList) As Integer

        Dim CurRegNode As String
        Dim NumNodesFound As Integer = 0

        CurRegNode = sBaseSW32 & NodeToTest
        For i As Integer = 0 To 1
            If RegKeyExists(RegistryHive.LocalMachine, CurRegNode) Then
                ReturnedRegNodes.Add(CurRegNode)
                NumNodesFound = NumNodesFound + 1
            End If
            CurRegNode = sBaseSW64 & NodeToTest
        Next

    End Function

    Private Function SearchPathNames(ByRef RegNode As String, ByRef ValueFound As String) As Boolean

        Dim PathFound As Boolean = False

        'search path names first, if not found - try default
        For Each ValName As String In My.Settings.PathKeyNames
            ValueFound = ValName
            If GetKeyValue(RegistryHive.LocalMachine, RegNode, ValueFound) Then
                If Not ValueFound = "" Then
                    If ValueFound.EndsWith(";") Then
                        ValueFound = ValueFound.Remove(ValueFound.Length - 1)
                    End If
                    PathFound = True
                    Exit For
                End If
            End If
        Next
        If Not PathFound Then
            ValueFound = ""
            If GetKeyValue(RegistryHive.LocalMachine, RegNode, ValueFound) Then
                If Not ValueFound = "" Then
                    If ValueFound.EndsWith(";") Then
                        ValueFound = ValueFound.Remove(ValueFound.Length - 1)
                    End If
                    PathFound = True
                End If
            End If
        End If
        Return PathFound

    End Function

    Private Function SearchRevNames(ByRef RegNode As String, ByRef ValueFound As String) As Boolean

        Dim VersionFound As Boolean

        VersionFound = False
        For Each ValName As String In My.Settings.VersionKeyNames
            ValueFound = ValName
            If GetKeyValue(RegistryHive.LocalMachine, RegNode, ValueFound) Then
                If Not ValueFound = "" Then
                    VersionFound = True
                    Exit For
                End If
            End If
        Next
        Return VersionFound

    End Function

    Private Function SearchInstallerPaths(ByVal ProductName As String, ByRef PathsFound As ArrayList) As Boolean

        Dim ReturnedValueNames As New ArrayList
        Dim InstallerPathsNode As String, ValueName As String
        Dim InstPath As String, PathParams() As String
        Dim CurPath As String

        SearchInstallerPaths = False
        InstallerPathsNode = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Folders"
        GetAllValues(RegistryHive.LocalMachine, InstallerPathsNode, ReturnedValueNames)
        For Each InstPath In My.Settings.PathsFromInstaller
            If InstPath.StartsWith(ProductName) Then
                PathParams = InstPath.Split("|")
                For Each ValueName In ReturnedValueNames
                    If ValueName.EndsWith(PathParams(1)) Then
                        CurPath = ValueName
                        PathsFound.Add(CurPath)
                        SearchInstallerPaths = True
                    End If
                Next
            End If
        Next


    End Function

    Private Function SearchForProductCode(ByRef RegNode As String, ByRef _
        CurProgram As String, ByRef CodeFound As String) As Boolean

        Dim WinUninstallNode As String, CurSubNode As String
        Dim ReturnedKeyNames As New ArrayList
        Dim NumNodesFound As Integer, CurKeyName As String
        Dim ValueFound As String, FoundProdCode As Boolean

        FoundProdCode = False
        ValueFound = "ProductCode"
        If GetKeyValue(RegistryHive.LocalMachine, RegNode, ValueFound) Then
            If Not ValueFound = "" Then
                CodeFound = ValueFound
                FoundProdCode = True
            End If
        End If
        If Not FoundProdCode Then
            WinUninstallNode = sBaseSW32 & sWinUInst
            NumNodesFound = FindSubNodes(RegistryHive.LocalMachine, WinUninstallNode, ReturnedKeyNames)
            For Each CurKeyName In ReturnedKeyNames
                CurSubNode = WinUninstallNode & "\" & CurKeyName
                ValueFound = "DisplayName"
                If GetKeyValue(RegistryHive.LocalMachine, CurSubNode, ValueFound) Then
                    If Not IsNothing(ValueFound) Then
                        If ValueFound.StartsWith(CurProgram) Then
                            CodeFound = CurKeyName
                            FoundProdCode = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        If Not FoundProdCode Then
            WinUninstallNode = sBaseSW64 & sWinUInst
            NumNodesFound = FindSubNodes(RegistryHive.LocalMachine, WinUninstallNode, ReturnedKeyNames)
            For Each CurKeyName In ReturnedKeyNames
                CurSubNode = WinUninstallNode & "\" & CurKeyName
                ValueFound = "DisplayName"
                If GetKeyValue(RegistryHive.LocalMachine, CurSubNode, ValueFound) Then
                    If Not IsNothing(ValueFound) Then
                        If ValueFound.StartsWith(CurProgram) Then
                            CodeFound = CurKeyName
                            FoundProdCode = True
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
        Return FoundProdCode

    End Function

    Private Function SearchUserDirs(ByRef PublicDirs As ArrayList) As Boolean

        Dim UserGroupDir() As String = Nothing
        Dim DirsFound As Boolean = False

        For Each gn As String In cProgGroupNodes.AllKeys
            Dim TrimedName As String = gn.Replace(" (32)", "")
            TrimedName = TrimedName.Replace(" (64)", "")
            If Directory.Exists(sPublicDocPath & TrimedName) Then
                UserGroupDir = _
                    Directory.GetDirectories(sPublicDocPath & TrimedName)
                DirsFound = True
            End If
        Next
        If DirsFound Then PublicDirs.AddRange(UserGroupDir)
        Return DirsFound

    End Function

    Private Function SearchForDrivers(ByRef DriversFound As ArrayList) As Boolean

        Dim ReturnedKeyNames As New ArrayList
        Dim DriverDirec As String
        Dim DirectoryList As New ArrayList
        Dim Success As Boolean = False
        Dim NameFound As Boolean

        Dim SysFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
        Dim FolderOpts() As String = {"\DRVSTORE", "\DriverStore", "\DriverStore\FileRepository"}
        Dim NamingOpts() As String = {""}
        Dim DriverCats As Collections.Specialized.StringCollection
        Dim RegisteredDrvr As String

        DriverCats = GetSelectedPackage("Drivers")
        'Select Case SelectedCat
        '    Case 0
        '        DriverCats = My.Settings.MCCDriverCategory
        '        ReDim NamingOpts(1)
        '        NamingOpts.SetValue("_", 0)
        '        NamingOpts.SetValue(".inf", 1)
        '    Case 1
        '        ReDim NamingOpts(0)
        '        DriverCats = My.Settings.DTDriverCategory
        '        NamingOpts.SetValue("", 0)
        '    Case Else
        'End Select
        If DriversFound.Contains("Stored") Then
            DriversFound.Clear()
            For OptNum As Integer = 0 To FolderOpts.Length - 1
                DriverDirec = SysFolder & FolderOpts(OptNum)
                If Directory.Exists(DriverDirec) Then
                    'GetDirectories(DriverDirec, DirectoryList)
                    Dim DirStrArray() As String = Directory.GetDirectories(DriverDirec)
                    DirectoryList.AddRange(DirStrArray)
                    For Each dn As String In DriverCats
                        For Each DriverDir As String In DirectoryList
                            For NameOpt As Integer = 0 To NamingOpts.Length - 1
                                Dim Opt As String = NamingOpts(NameOpt)
                                NameFound = DriverDir.Contains(dn & Opt)
                                If NameFound Then
                                    DriversFound.Add(DriverDir)
                                    Success = True
                                End If
                            Next
                        Next
                    Next
                End If
            Next
        End If
        If DriversFound.Contains("Registered") Then
            DriversFound.Clear()
            Dim RevDrvs As String = "SYSTEM\CurrentControlSet\Services"
            Dim NodesFound As Integer = FindSubNodes(RegistryHive.LocalMachine, RevDrvs, ReturnedKeyNames)
            For Each rkn As String In ReturnedKeyNames
                For Each MCDriver As String In DriverCats
                    RegisteredDrvr = rkn.ToLower
                    If RegisteredDrvr.Contains(MCDriver) Then
                        Dim ValName As String = "ImagePath"
                        If GetKeyValue(RegistryHive.LocalMachine, RevDrvs & "\" & rkn, ValName) Then
                            Dim ip As String = ValName
                            Dim WinPath As String = Environment.GetEnvironmentVariable("WINDIR")
                            Dim DrvFileName As String = WinPath & "\" & ValName
                            If File.Exists(DrvFileName) Then
                                DriversFound.Add(DrvFileName)
                                Success = True
                            End If
                        End If
                    End If
                Next
            Next
        End If
        Return Success

    End Function

    Private Function SearchForPlatforms(ByRef PlatformsFound As ArrayList) As Boolean

        Dim ReturnedKeyNames As New ArrayList
        Dim PFound As Boolean = False

        For Each PlfmDir As String In My.Settings.Platforms
            ReturnedKeyNames.Clear()
            Dim NodesFound As Integer = FindSubNodes(RegistryHive.LocalMachine, PlfmDir, ReturnedKeyNames)
            For Each rrn As String In ReturnedKeyNames
                Dim sln As Integer = rrn.Length - 1
                If sln > 3 Then sln = 3
                Dim fwv As String = rrn.Substring(1, sln)
                Dim ValName As String = "Version"
                Dim ServPack As String = ""
                Dim fwType As String = "  Framework "
                If PlfmDir.Contains("Compact") Then fwType = "  Compact Framework "
                If GetKeyValue(RegistryHive.LocalMachine, PlfmDir & "\" & rrn, ValName) Then
                    fwv = ValName
                Else
                    Dim NewNode As String = PlfmDir & "\" & rrn '& "\Client"
                    Dim ReturnedNodes As New ArrayList
                    Dim ClNodeFound As Integer = FindSubNodes(RegistryHive.LocalMachine, NewNode, ReturnedNodes)
                    For Each rns As String In ReturnedNodes
                        If rns = "Client" Then
                            ValName = "Version"
                            If GetKeyValue(RegistryHive.LocalMachine, NewNode & "\" & rns, ValName) Then
                                fwv = ValName
                            End If
                        End If
                    Next
                End If
                ValName = "SP"
                ServPack = ""
                If GetKeyValue(RegistryHive.LocalMachine, PlfmDir & "\" & rrn, ValName) Then
                    If Not ValName.StartsWith("0") Then ServPack = " SP " & ValName
                End If
                PFound = True
                PlatformsFound.Add(fwType & fwv & ServPack)
                sbContents.AppendLine(fwType & fwv & ServPack)
            Next
        Next
        Return PFound

    End Function

    Private Sub GetDirectories(ByVal StartPath As String, ByRef DirectoryList As ArrayList)

        Dim Dirs() As String = Directory.GetDirectories(StartPath)
        DirectoryList.AddRange(Dirs)
        For Each Dir As String In Dirs
            GetDirectories(Dir, DirectoryList)
        Next

    End Sub

    Private Sub CreateDirListing(ByVal CurPath As String, ByVal Filter As String, Optional ByVal GetVolume As Boolean = False)

        Dim FilesFound As String()
        Dim SplitSet() As String
        Dim DirList As New ArrayList
        Dim fi As FileInfo
        Dim FileDT As Date, NumFiles As Integer
        Dim NameOfFile As String, SizeOfFile As String
        Dim DirHeading, UserName, Filler As String
        Dim ListGroup As Boolean = False
        Dim DirIndicator As String = "    <DIR>          "

        SplitSet = Filter.Split("|")
        DirHeading = ""
        Filler = ""
        If Filter.StartsWith("MCCMenu") Then
            Filter = "*.*"
            Filler = "  "
            For Each ProgGroup As String In My.Settings.MCCMenu
                If CurPath.Contains(ProgGroup) Then
                    ListGroup = True
                    Dim PathSplit As Integer
                    If CurPath.Contains("Programs") Then
                        PathSplit = CurPath.IndexOf("Programs")
                    Else
                        PathSplit = CurPath.IndexOf(ProgGroup)
                    End If
                    If SplitSet.Length = 1 Then
                        DirHeading = "Menu for All Users: " & vbCrLf & _
                            Filler & CurPath.Substring(PathSplit) & vbCrLf
                    Else
                        For Each UserName In SplitSet
                            If CurPath.Contains(UserName) Then
                                DirHeading = "Menu for " & UserName & ":" & vbCrLf & _
                                    Filler & CurPath.Substring(PathSplit) & vbCrLf
                            End If
                        Next
                    End If
                    Exit For
                End If
            Next
        Else
            Dim Disk As System.Management.ManagementObject

            ListGroup = True
            If GetVolume Then
                Dim volPhrase As String
                Dim di As New System.IO.DriveInfo(CurPath)
                Dim volName As String = di.VolumeLabel
                Dim win32Drive As String = _
                    "Win32_LogicalDisk='" & CurPath.Substring(0, 2) & "'"
                Disk = New System.Management.ManagementObject(win32Drive)
                Dim volSer As String = Disk.Properties.Item("VolumeSerialNumber").Value
                If volName = "" Then
                    volPhrase = " has no label."
                Else
                    volPhrase = " is " & volName
                End If
                DirHeading = " Volume in drive " & CurPath.Substring(0, 1) _
                    & volPhrase & vbCrLf & " Volume Serial Number" _
                    & " is " & volSer & vbCrLf & vbCrLf & " Directory of " _
                    & CurPath & vbCrLf
            Else
                DirHeading = " Directory of " & _
                    CurPath & vbCrLf
            End If
        End If

        If ListGroup Then
            Dim FileCount, AggregateSize As Long
            FilesFound = Directory.GetFileSystemEntries(CurPath, Filter)
            Dim rt As DirectoryInfo = Directory.GetParent(CurPath)
            'Dim rrt As DirectoryInfo = rt.Parent

            NumFiles = FilesFound.Length
            If NumFiles > 0 Then sbContents.AppendLine(DirHeading)
            sbContents.AppendLine(Directory.GetLastWriteTime(CurPath).ToString _
                ("MM/dd/yyyy  hh:mm tt") & DirIndicator & ".")
            If Not IsNothing(rt) Then sbContents.AppendLine _
                (rt.LastWriteTime.ToString("MM/dd/yyyy  " & _
                "hh:mm tt") & DirIndicator & "..")

            For CurFile As Integer = 0 To NumFiles - 1
                fi = New FileInfo(FilesFound(CurFile))
                NameOfFile = fi.Name
                FileDT = fi.LastWriteTime
                If fi.Attributes And FileAttributes.Directory Then
                    sbContents.AppendLine(Filler & FileDT.ToString("MM/dd/yyyy  hh:mm tt") _
                        & DirIndicator & NameOfFile)
                Else
                    SizeOfFile = fi.Length.ToString("0,0")
                    sbContents.AppendLine(Filler & FileDT.ToString("MM/dd/yyyy  hh:mm tt") _
                        & vbTab & SizeOfFile.PadLeft(14) & " " & NameOfFile)
                    FileCount = FileCount + 1
                    AggregateSize = AggregateSize + fi.Length
                End If
            Next
            sbContents.AppendLine(vbTab & "      " & FileCount.ToString _
                & " File(s)     " & AggregateSize.ToString("0,0") _
                & " bytes" & vbCrLf)
        End If

    End Sub

    Private Sub CreatePropsListing(ByVal CurPath As String, ByVal Filter As String, _
        Optional ByVal ProgSelected As String = "", Optional ByVal ShowGUID As Boolean = True)

        Dim FilesFound As String() = {CurPath}
        Dim DirList As New ArrayList
        Dim fi As FileInfo
        Dim fvi As FileVersionInfo
        Dim FileDT As Date, NumFiles As Integer
        Dim NameOfFile As String, VerOfFile As String = ""
        Dim VerOfProd As String = "", ShortName As String
        Dim Head0 As String = ""
        Dim Head1 As String = "File Version"
        Dim Head2 As String = "Program Version"
        Dim Head3 As String = "File Date"
        Dim ColHead As String = ""
        Dim GACCats As Collections.Specialized.StringCollection
        Dim ListGroup As Boolean, ListGac As Boolean
        Dim ValidProd As Boolean
        Dim PrintHeader As Boolean = True
        Dim PrintSubHeader As Boolean = True
        Dim DriverListing As Boolean = False

        'GACCats = My.Settings.GACEntries
        GACCats = GetSelectedPackage("GAC")
        ListGac = False
        'ShowGUID = True
        If ProgSelected = "Drivers" Then
            DriverListing = True
            'sDvrList
            ProgSelected = ""
        End If
        If Filter = "MCCGac" Then
            Filter = "*.*"
            ListGac = True
            ProgSelected = "Global Assemblies"
            ValidProd = True
        Else
            If ProgSelected = "" Then
                ProgSelected = Me.lstValues.SelectedItem
            End If
            ValidProd = Not (ProgSelected = "")
        End If

        If CurPath.StartsWith(sSystemPath) Then
            Head0 = "Shared Files for " & ProgSelected
            ShowGUID = False
            PrintHeader = False
            If ValidProd Then sbContents.AppendLine(Head0)
            sbContents.AppendLine(CurPath & vbCrLf)
        ElseIf CurPath.StartsWith("SharedFiles") Then
            CurPath = sSystemPath
            ShowGUID = False
            PrintHeader = False
        Else
            If ValidProd And (sbContents.Length = 0) Then Head0 = "File properties for " & ProgSelected
        End If
        If Directory.Exists(CurPath) Then
            FilesFound = Directory.GetFileSystemEntries(CurPath, Filter)
        End If
        NumFiles = FilesFound.Length

        ColHead = "File" & Head1.PadLeft(37) & _
                Head2.PadLeft(20) & Head3.PadLeft(12)
        For CurFile As Integer = 0 To NumFiles - 1
            If PrintHeader And Not ListGac Then
                If ValidProd Then sbContents.AppendLine(Head0)
                sbContents.AppendLine(CurPath)
                If ValidProd And ShowGUID Then
                    sbContents.AppendLine("Package Version GUID:  " & _
                        cMCCProgGUIDs.Get(ProgSelected))
                End If
                sbContents.AppendLine()
                sbContents.AppendLine(ColHead)
                sbContents.AppendLine()
                PrintHeader = False
            End If
            fi = New FileInfo(FilesFound(CurFile))
            NameOfFile = fi.Name
            If DriverListing And NameOfFile.EndsWith(".inf") Then
                sDvrList.Add(NameOfFile, CurPath)
            End If
            If (Not (fi.Attributes And FileAttributes.Directory) = FileAttributes.Directory) Then
                ListGroup = False
                If ListGac Then
                    For Each ProgGroup As String In GACCats
                        If NameOfFile.StartsWith(ProgGroup) Then
                            ListGroup = True
                            If PrintHeader Then
                                If ValidProd And PrintSubHeader Then sbContents.AppendLine(Head0)
                                sbContents.AppendLine(CurPath)
                                sbContents.AppendLine()
                                sbContents.AppendLine(ColHead)
                                sbContents.AppendLine()
                                PrintHeader = False
                                PrintSubHeader = False
                            End If
                            Exit For
                        End If
                    Next
                Else
                    ListGroup = True
                End If
                If ListGroup Then
                    fvi = FileVersionInfo.GetVersionInfo(FilesFound(CurFile))
                    VerOfFile = ""
                    VerOfProd = ""
                    Dim fVersion As String = fvi.FileVersion
                    Dim ProdVer As String = fvi.ProductVersion
                    If Not IsNothing(fVersion) Then
                        Dim FileVer As String = fvi.FileMajorPart & _
                            "." & fvi.FileMinorPart & "." & _
                            fvi.FileBuildPart & "." & _
                            fvi.FilePrivatePart
                        VerOfFile = FileVer.PadRight(15)
                    Else
                        VerOfFile = VerOfFile.PadRight(15)
                    End If
                    If Not IsNothing(ProdVer) Then VerOfProd = ProdVer.PadRight(15)
                    FileDT = fi.LastWriteTime
                    ShortName = NameOfFile
                    If NameOfFile.Length > 28 Then
                        ShortName = NameOfFile.Remove(4, NameOfFile.Length - 28)
                        ShortName = ShortName.Insert(4, "_")
                    End If
                    sbContents.AppendLine("  " & ShortName.PadRight(26) & vbTab & _
                        VerOfFile.PadRight(15) & vbTab & VerOfProd.PadRight(15) & _
                        vbTab & FileDT.ToString("MM/dd/yyyy  hh:mm tt"))
                End If
            End If
        Next
        If ListGroup And (NumFiles > 0) And ShowGUID Then sbContents.AppendLine()

    End Sub

    Private Function CheckForCustomRegNodes(ByRef CurProgName As String, ByRef RegNodes As ArrayList) As Integer

        Dim CurRegNode As String = Nothing
        Dim CustomRegNode As String = Nothing
        Dim SplitSet() As String, ReturnedKeyNames As New ArrayList
        Dim ReturnedRegNodes As New ArrayList
        Dim PathFound As Boolean, NumSubNodes As Integer
        Dim NumRegNodes As Integer, AlreadyFound As Boolean

        For Each StoredRegNode As String In My.Settings.CustomRegNodes
            SplitSet = StoredRegNode.Split("|")
            If SplitSet(0) = CurProgName Then
                CustomRegNode = SplitSet(1)
                Dim NodesFound As Integer = GetRegNodes(CustomRegNode, ReturnedRegNodes)
                For Each Node As String In ReturnedRegNodes
                    PathFound = SearchPathNames(Node, CurRegNode)
                    If PathFound Then
                        For Each s As String In RegNodes
                            If s = CurRegNode Then AlreadyFound = True
                        Next
                        If Not AlreadyFound Then
                            RegNodes.Add(CurRegNode)
                            NumRegNodes = NumRegNodes + 1
                        End If
                        AlreadyFound = False
                    Else
                        NumSubNodes = FindSubNodes(RegistryHive.LocalMachine, Node, ReturnedKeyNames)
                        For Each rkn As String In ReturnedKeyNames
                            PathFound = False
                            Dim NewRegNode As String = Node & "\" & rkn
                            PathFound = SearchPathNames(NewRegNode, CurRegNode)
                            If PathFound Then
                                For Each s As String In RegNodes
                                    If s = CurRegNode Then AlreadyFound = True
                                Next
                                If Not AlreadyFound Then
                                    RegNodes.Add(CurRegNode)
                                    NumRegNodes = NumRegNodes + 1
                                End If
                            End If
                            AlreadyFound = False
                        Next
                    End If
                Next
            End If
        Next
        Return NumRegNodes

    End Function

    Private Function CheckForCustomSubDirs(ByRef CurProgName As String, ByRef SubDirs As ArrayList) As Integer

        Dim SplitSet() As String, VarSplit() As String
        'Dim SplitGroup() As String
        Dim NumSubsFound As Integer

        For Each sd As String In My.Settings.CustomSubDirs
            'SplitGroup = sd.Split("@")
            'If SplitGroup.Length > 1 Then 
            'add custom (TracerDAQ|%APPDATA%@ProgGroups) for app data
            SplitSet = sd.Split("|")
            If SplitSet(0) = CurProgName Then
                Dim TempDir As String = SplitSet(1)
                Dim TempLoc As Integer = TempDir.LastIndexOf("%")
                If TempLoc > 0 Then
                    VarSplit = TempDir.Split("%")
                    Dim SysEnvDir As String = VarSplit(1)
                    Dim SysEnvPath As String = Environment.GetEnvironmentVariable(SysEnvDir)
                    TempDir = SysEnvPath & VarSplit(2)
                End If
                SubDirs.Add(TempDir)
                NumSubsFound = NumSubsFound + 1
            End If
        Next
        Return NumSubsFound

    End Function

    Private Function GetCurOS(ByRef OSVersion As String, ByRef OSShort As String) As String

        Dim OSBitsShort As String = "-32"
        Dim OSString As String = "", OSStringShort As String = ""

        sOSBits = "32-bit"
        OSVersion = osVer.ToString
        Select Case osPltfm
            Case PlatformID.Win32NT
                Select Case osVer.Major
                    Case 10
                        OSString = "Windows 10"
                        OSStringShort = "W10"
                        Dim ValueFound As String = "CurrentBuild"
                        Dim RegNode As String = _
                            "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                        If RegKeyExists(RegistryHive.LocalMachine, RegNode) Then
                            Dim GotKeyVal As Boolean = GetKeyValue(RegistryHive.LocalMachine, _
                                RegNode, ValueFound)
                            If Not IsNothing(ValueFound) Then
                                If Not ValueFound = "" Then
                                    Dim MinorBuild As String = "UBR"
                                    Dim GotSubVal As Boolean = GetKeyValue _
                                        (RegistryHive.LocalMachine, RegNode, MinorBuild)
                                    Dim MB As String = ""
                                    If Not (MinorBuild = "") Then MB = "." & MinorBuild
                                    OSString = "Windows 10 build " & ValueFound & MB
                                    OSStringShort = "W10"
                                End If
                            End If
                        End If
                    Case 6, Is > 6
                        Select Case osVer.Minor
                            Case 1
                                OSString = "Windows 7"
                                OSStringShort = "W7"
                            Case 2
                                OSString = "Windows 8"
                                OSStringShort = "W8"
                                Dim ValueFound As String = "CurrentVersion"
                                Dim RegNode As String = _
                                    "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
                                If RegKeyExists(RegistryHive.LocalMachine, RegNode) Then
                                    Dim GotKeyVal As Boolean = GetKeyValue(RegistryHive.LocalMachine, _
                                        RegNode, ValueFound)
                                    If Not IsNothing(ValueFound) Then
                                        If Not ValueFound = "" Then
                                            If ValueFound.StartsWith("6.3") Then
                                                OSString = "Windows 8.1"
                                                OSStringShort = "W8.1"
                                            End If
                                            If ValueFound.StartsWith("6.4") Then
                                                OSString = "Windows 10"
                                                OSStringShort = "W10"
                                            End If
                                        End If
                                    End If
                                End If
                            Case Else
                                OSString = "Windows Vista"
                                OSStringShort = "Vista"
                        End Select
                    Case 5
                        Select Case osVer.Minor
                            Case 0
                                OSString = "Windows 2000"
                                OSStringShort = "W2k"
                            Case Else
                                OSString = "Windows XP"
                                OSStringShort = "XP"
                        End Select
                    Case Else
                        OSString = "Windows NT"
                        OSStringShort = "NT"
                End Select
            Case PlatformID.Win32Windows
                OSString = "Windows ME, 98, or 95"
                OSStringShort = "W9x"
            Case PlatformID.Win32S
                OSString = "Really Old Windows"
                OSStringShort = "W3"
        End Select

        Dim pa As String = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")
        If pa.Contains("64") Then
            sOSBits = "64-bit"
            OSBitsShort = "-64"
        End If
        OSShort = OSStringShort & OSBitsShort
        Return OSString & " (" & sOSBits & ")  "

    End Function

    Private Function GetPCName(ByRef ShortName As String) As String

        Dim NameFound As String
        Dim LengthFound As Integer = 1

        NameFound = Environment.GetEnvironmentVariable("COMPUTERNAME")
        If Not IsNothing(NameFound) Then
            If NameFound.Length > 4 Then
                LengthFound = NameFound.Length - 4
            End If
        End If
        ShortName = NameFound.Substring(LengthFound)
        Return NameFound

    End Function

    Private Function GetDevelopmentEnvirons() As Boolean

        Dim ReturnedRegNodes As New ArrayList
        Dim NodesReturned As Integer, ValueFound As String = ""
        Dim ReturnedKeyNames As New ArrayList
        Dim FoundDE As Boolean = False
        Dim NodeSplit() As String
        Dim RevFound As String = ""

        For Each de As String In My.Settings.DevelopmentProgs
            NodeSplit = de.Split("|")
            Dim ProdName As String = NodeSplit(0)
            Dim ProdNode As String = NodeSplit(1)
            NodesReturned = GetRegNodes(ProdNode, ReturnedRegNodes)
            For Each rrn As String In ReturnedRegNodes
                ReturnedKeyNames.Clear()
                RevFound = ""
                NodesReturned = FindSubNodes(RegistryHive.LocalMachine, rrn, ReturnedKeyNames)
                If NodesReturned = 0 Then
                    Dim CurKeyVal As String = rrn
                    Dim Found As Boolean = SearchPathNames(CurKeyVal, ValueFound)
                    If Found Then
                        If Directory.Exists(ValueFound) Then
                            If (ProdName = "Android Studio") Then
                                sbContents.AppendLine("  " & ProdName)
                                FoundDE = True
                            End If
                        End If
                    End If
                Else
                    For Each rkn As String In ReturnedKeyNames
                        Dim CurKeyVal As String = rrn & "\" & rkn
                        Dim Found As Boolean = SearchPathNames(CurKeyVal, ValueFound)
                        If (ProdName = "Visual Studio") Then
                            If (Not Found) Then
                                Dim NewKeyVal = rrn & "\" & rkn & "\Setup"
                                Found = SearchPathNames(NewKeyVal, ValueFound)
                            End If
                            RevFound = "  " & rkn
                        End If
                        If Found Then
                            If Directory.Exists(ValueFound) Then
                                Dim AppName As String = GetAppName(ProdName, CurKeyVal)
                                If AppName.Length > 0 Then
                                    Dim ProgBits As String = ""
                                    If ProdName = "LabVIEW" Then
                                        Dim ParentKey As String = _
                                            "SOFTWARE\National Instruments\Common\Installer"
                                        Dim ProgDir As String = "NIDIR64"
                                        If GetKeyValue(RegistryHive.LocalMachine, ParentKey, ProgDir) Then
                                            If ValueFound.Contains(ProgDir) Then _
                                                ProgBits = " 64-bit"
                                        End If
                                    End If
                                    Dim RevReturned As String = ""
                                    If SearchRevNames(CurKeyVal, RevReturned) Then
                                        RevFound = "  ver " & RevReturned
                                    End If
                                    sbContents.AppendLine("  " & AppName & ProgBits & RevFound)
                                    FoundDE = True
                                End If
                            End If
                        End If
                        Found = False
                    Next
                End If
            Next
            ReturnedRegNodes.Clear()
        Next
        Return FoundDE

    End Function

    Private Function GetDependencies() As Boolean

        Dim ReturnedRegNodes As New ArrayList
        Dim NodesReturned As Integer, ValueFound As String = ""
        Dim ReturnedKeyNames As New ArrayList
        Dim FoundDE As Boolean = False
        Dim NodeSplit() As String
        Dim RevFound As String = ""
        Dim Found As Boolean

        For Each de As String In My.Settings.Dependencies
            NodeSplit = de.Split("|")
            Dim DepProdName As String = NodeSplit(0)
            Dim DepProdNode As String = NodeSplit(1)
            NodesReturned = GetRegNodes(DepProdNode, ReturnedRegNodes)
            For Each rrn As String In ReturnedRegNodes
                ReturnedKeyNames.Clear()
                RevFound = ""
                Found = False
                NodesReturned = FindSubNodes(RegistryHive.LocalMachine, rrn, ReturnedKeyNames)
                For Each rkn As String In ReturnedKeyNames
                    Dim CurKeyVal As String = rrn & "\" & rkn
                    Found = SearchPathNames(CurKeyVal, ValueFound)
                    If Found Then
                        If Directory.Exists(ValueFound) Then
                            Dim AppName As String = GetAppName(rkn, CurKeyVal)
                            If AppName.Length > 0 Then
                                Dim ProgBits As String = ""
                                Dim RevReturned As String = ""
                                If SearchRevNames(CurKeyVal, RevReturned) Then
                                    RevFound = "  ver " & RevReturned
                                End If
                                sbContents.AppendLine("  " & AppName & ProgBits & RevFound)
                                FoundDE = True
                                Found = False
                                Exit For
                            End If
                        End If
                    End If
                    Found = False
                Next
                If Not Found Then
                    Dim CustomNodesReturned As New ArrayList
                    For Each rn As String In My.Settings.CustomRegNodes
                        NodeSplit = rn.Split("|")
                        Dim ProdName As String = NodeSplit(0)
                        Dim ProdNode As String = NodeSplit(1)
                        If (DepProdName = ProdName) Then
                            If (ProdNode = "NONE") Then
                                Found = True
                            Else
                                Found = SearchPathNames(ProdNode, ValueFound)
                                If Not Found Then
                                    Found = SearchPathNames("SOFTWARE\" & ProdNode, ValueFound)
                                End If
                            End If
                            If Found Then
                                If Directory.Exists(ValueFound) Then
                                    Dim ProgBits As String = ""
                                    Dim RevReturned As String = ""
                                    If SearchRevNames(rrn, RevReturned) Then
                                        RevFound = "  ver " & RevReturned
                                    End If
                                    sbContents.AppendLine("  " & ProdName & ProgBits & RevFound)
                                    FoundDE = True
                                End If
                            End If
                            Found = False
                        End If
                    Next
                End If
            Next
            ReturnedRegNodes.Clear()
        Next
        Return FoundDE

    End Function

    Private Function GetOtherEntries() As Boolean

        Dim ReturnedRegNodes As New ArrayList
        Dim NodesReturned As Integer, ValueFound As String = ""
        'Dim ReturnedKeyNames As New ArrayList
        Dim FoundOE As Boolean = False
        Dim NodeSplit() As String
        Dim RevFound As String = ""

        For Each ent As String In My.Settings.OtherMCCEntries
            NodeSplit = ent.Split("|")
            Dim ProdName As String = NodeSplit(0)
            Dim ProdNode As String = NodeSplit(1)
            NodesReturned = GetRegNodes(ProdNode, ReturnedRegNodes)
            For Each rrn As String In ReturnedRegNodes
                FoundOE = True
                ValueFound = ""
                sbContents.AppendLine("  " & ProdName)
                If GetKeyValue(RegistryHive.LocalMachine, rrn, ValueFound) Then
                    sbContents.AppendLine("     (Path = " & ValueFound & ")")
                    If ProdName = "DT Framework ReferenceEx" Then
                        cOddRegNodes.Add("DotNet", rrn)
                        'strip down this path to get base path
                        Dim loc As Integer = ValueFound.LastIndexOf("\")
                        Dim TrimPath As String = ValueFound.Remove(ValueFound.Length - 1)
                        For index As Integer = 0 To 2
                            loc = TrimPath.LastIndexOf("\")
                            TrimPath = TrimPath.Remove(loc)
                        Next
                        cOddRegNodes.Add("Omni CD", TrimPath)
                    End If
                End If
            Next
            ReturnedRegNodes.Clear()
        Next
        Return FoundOE

    End Function

    Private Function GetAppName(ByRef AppKey As String, ByRef AppRegNode As String) As String

        Dim ProdNameFound As String = ""
        Dim ReturnedKeyNames As New ArrayList
        Dim NodesReturned As Integer

        For Each pnk As String In My.Settings.ProdNameKeys
            If pnk.EndsWith("\") Then
                'search subkeys, not values
                Dim SearchKey As String = AppRegNode & "\" & pnk
                NodesReturned = FindSubNodes(RegistryHive.LocalMachine, SearchKey, ReturnedKeyNames)
                For Each Node As String In ReturnedKeyNames
                    If Node.Contains(AppKey) Then
                        ProdNameFound = Node
                        Exit For
                    End If
                Next
            Else
                'search value name
                If GetKeyValue(RegistryHive.LocalMachine, AppRegNode, pnk) Then
                    ProdNameFound = pnk
                    Exit For
                End If
            End If
        Next
        Return ProdNameFound

    End Function

    Private Sub FileLister_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim VolSplit() As String, CurVol As String = "CurVolume"
        Dim OSText As String, VersionString As String = ""
        Dim OSShort As String = "", ShortName As String = ""
        Dim SizeParams As String, LocParams As String
        Dim RegNode As String, ValSet As Boolean
        Dim L As String, T As String, W As String, H As String

        RegNode = "Software\Measurement Computing\FileLister"
        SizeParams = "Size"
        LocParams = "Location"
        ValSet = GetKeyValue(RegistryHive.CurrentUser, RegNode, SizeParams)
        ValSet = GetKeyValue(RegistryHive.CurrentUser, RegNode, LocParams)
        H = SizeParams.Remove(SizeParams.IndexOf(","))
        W = SizeParams.Remove(0, SizeParams.IndexOf(",") + 1)
        T = LocParams.Remove(LocParams.IndexOf(","))
        L = LocParams.Remove(0, LocParams.IndexOf(",") + 1)
        Me.Height = Val(H)
        Me.Width = Val(W)
        Me.Top = Val(T)
        Me.Left = Val(L)

        OSText = GetCurOS(VersionString, OSShort)
        Dim PCName As String = GetPCName(ShortName)

        Dim RegVolExists As Boolean = GetKeyValue( _
            RegistryHive.CurrentUser, sVolumeRegKey, CurVol)
        If RegVolExists Then
            Me.txtOutput.Text = "Current Volume = " & CurVol
        Else
            CurVol = ""
            Dim FLConfigPath As String = Environment.GetEnvironmentVariable("WINDIR") & "\FileListCfg.ini"
            If Not IO.File.Exists(FLConfigPath) Then FLConfigPath = _
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & _
                "\VirtualStore\Windows\FileListCfg.ini"
            If IO.File.Exists(FLConfigPath) Then
                Dim CfgText As String = File.ReadAllText(FLConfigPath)
                Me.txtOutput.Text = CfgText
                VolSplit = CfgText.Split("=")
                If Not IsNothing(VolSplit) Then
                    Dim TrimChars() As Char = {Chr(10), Chr(13)}
                    If VolSplit.Length = 2 Then CurVol = _
                        VolSplit(1).Trim(TrimChars) & " "
                End If
            End If
        End If
        Dim AndroidIsUnzipped As Boolean = UnzippedAndroidDirExists()
        If AndroidIsUnzipped Then
            'cmdUnZip.Enabled = True
            cmdUnZip.Text = "Del Archive"
        End If
        sCurVolume = CurVol & " " & OSShort & " " & ShortName

    End Sub

    Private Sub FileLister_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing

        Dim SizeSet As Boolean

        If Me.WindowState = FormWindowState.Normal Then
            Dim RegNode As String = "Software\Measurement Computing\FileLister"
            Dim SizeParams As String
            Dim LocParams As String
            SizeParams = Format(Me.Height, "0") & ", " & Format(Me.Width, "0")
            LocParams = Format(Me.Top, "0") & ", " & Format(Me.Left, "0")
            SizeSet = SetKeyValue(RegistryHive.CurrentUser, RegNode, "Size", SizeParams)
            SizeSet = SetKeyValue(RegistryHive.CurrentUser, RegNode, "Location", LocParams)
        End If

    End Sub

    Private Sub FileLister_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        Me.txtOutput.Height = Me.Height - 296

    End Sub

    Private Function ParseFileFilter(ByRef PathParam As String, ByRef Filter As String) As String

        Dim TempString() As String
        Dim PathResult As String = ""

        Filter = "*.*"
        TempString = Split(PathParam, "?")
        PathResult = TempString(0)
        If TempString.Length > 1 Then Filter = TempString(1)
        Return PathResult

    End Function

    Private Function UnzippedAndroidDirExists() As Boolean

        Dim NumSubDirs As Integer
        Dim SubDirs As New ArrayList
        Dim CurProgName As String = "Instacal & Universal Library"
        Dim DirExists As Boolean

        DirExists = False
        NumSubDirs = CheckForCustomSubDirs(CurProgName, SubDirs)
        For Each DirFound As String In SubDirs
            If DirFound.Contains("AndroidLib") Then
                If Directory.Exists(DirFound) Then
                    sAndroidTempPath = DirFound
                    DirExists = True
                    Exit For
                End If
            End If
        Next
        Return DirExists

    End Function

    Private Function GetAndroidVersion() As String

        Dim AndroidVersion As String = ""
        If Not (sAndroidTempPath = "") Then
            If Directory.Exists(sAndroidTempPath & "\assets\props") Then
                Dim XMLFile As String = sAndroidTempPath & "\assets\props\ul_props.xml"
                If File.Exists(XMLFile) Then
                    Dim itemRead As String, attRead As String
                    Dim reader As XmlTextReader = New XmlTextReader(XMLFile)
                    attRead = ""
                    Do While (reader.Read())
                        Select Case reader.NodeType
                            Case XmlNodeType.Element
                                If reader.HasAttributes Then
                                    While reader.MoveToNextAttribute()
                                        itemRead = reader.Name
                                        attRead = reader.Value
                                    End While
                                End If
                            Case XmlNodeType.Text
                                itemRead = reader.Value
                                If attRead = "UL_VERSION" Then AndroidVersion = itemRead
                        End Select
                    Loop
                End If
            End If
        End If
        Return AndroidVersion

    End Function

    Private Sub cmdUnZip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUnZip.Click

        Dim PathString As String
        Dim PathParam As String, FilterString As String = "*.*"

        If lstValNames.SelectedIndex < 0 Then Exit Sub
        PathParam = lstValNames.SelectedItem.ToString
        PathString = ParseFileFilter(PathParam, FilterString)
        PathString = Chr(34) & PathString & "Android\ul.jar" & Chr(34)
        Dim BatchCmd As Process

        If sAndroidTempPath = "" Then
            BatchCmd = New Process
            BatchCmd.StartInfo.FileName = "UZCmd.bat"
            BatchCmd.StartInfo.Arguments = PathString & " AndroidLib"
            BatchCmd.Start()
            BatchCmd.WaitForExit()
            If UnzippedAndroidDirExists() Then
                cmdUnZip.Enabled = False
            End If
        Else
            Dim di As New System.IO.DirectoryInfo(sAndroidTempPath)
            Dim reslt As MsgBoxResult = MsgBox( _
                "Delete previously unzipped Android files?", _
                MsgBoxStyle.YesNo, "Delete Android Listing")
            If reslt = MsgBoxResult.Yes Then
                di.Delete(True)
                sAndroidTempPath = ""
                cmdUnZip.Text = "Unzip Files"
            End If
        End If
        cmdStart.Select()
        cmdStart_Click(Me.cmdUnZip, e)
        lstValues.Items.Clear()
        lstValNames.Items.Clear()
        cmdUnZip.Enabled = False

    End Sub

    Private Sub UpdateTextList(ByVal TargetFile As String, ByVal ClearPrevious As Boolean)

        Dim TextExists As Boolean
        Dim EnableCmds As Boolean

        If txtOutput.Text.StartsWith("[CurrentVersion]") _
            Or txtOutput.Text.StartsWith("Current Volume") _
            Then ClearPrevious = True
        If ClearPrevious Then
            txtOutput.Text = sbContents.ToString
        Else
            txtOutput.AppendText(sbContents.ToString)
        End If
        TextExists = (txtOutput.TextLength > 0)
        If TextExists Then
            cmdSave.Enabled = True
            cmdClear.Enabled = True
            If (txtFileName.Text = "") Or ClearPrevious _
                Then txtFileName.Text = TargetFile
        End If

        EnableCmds = Not (lstKeys.SelectedIndex < 0)
        cmdDriverList.Enabled = EnableCmds
        cmdListGac.Enabled = EnableCmds
        cmdStartList.Enabled = EnableCmds


    End Sub

    Private Sub DriversToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriversToolStripMenuItem.Click

        Dim DrvrCount As Integer = sDvrList.Count
        frmDrivers.ShowDialog(Me)
        If Not (DrvrCount = Me.sDvrList.Count) Then
            txtOutput.Clear()
            FindDrivers()
        End If
        frmDrivers.Close()

    End Sub

    Private Sub FindDrivers()

        Dim DriversFound As New ArrayList
        Dim DirList As New ArrayList
        Dim TargetFile As String
        Dim Filter As String = "*.*"

        sDvrList.Clear()
        cmdListFiles.Enabled = False
        cmdListProps.Enabled = False
        Me.lstValues.SelectedIndex = -1
        Me.lstValNames.Items.Clear()
        sbContents.Length = 0
        TargetFile = sDocPath & "\Drvr " & sCurVolume & ".txt"

        DriversFound.Add("Stored")
        Dim FoundDrivers As Boolean = SearchForDrivers(DriversFound)
        If FoundDrivers Then
            'If Not (txtFileName.Text = "") Then sbContents.AppendLine()
            sbContents.AppendLine("Currently stored drivers:" & vbCrLf)
            For Each df As String In DriversFound
                CreatePropsListing(df, Filter, "Drivers")
                sbContents.AppendLine()
            Next
        End If
        DriversFound.Clear()
        DriversFound.Add("Registered")
        FoundDrivers = SearchForDrivers(DriversFound)
        If FoundDrivers Then
            sbContents.AppendLine("Currently registered drivers:" & vbCrLf)
            For Each df As String In DriversFound
                CreatePropsListing(df, Filter)
                sbContents.AppendLine()
            Next
        End If
        Me.DriversToolStripMenuItem.Enabled = (sDvrList.Count > 0)

        UpdateTextList(TargetFile, False)

    End Sub

    Private Function GetSelectedPackage(ByVal CollectionType As String) As Collections.Specialized.StringCollection

        Dim CurProdGroup As String
        Dim SelectedCollection As System.Collections.Specialized.StringCollection

        SelectedCollection = My.Settings.MCCDriverCategory
        If Not lstKeys.SelectedIndex < 0 Then
            CurProdGroup = lstKeys.SelectedItem.ToString
            Select Case CollectionType
                Case "GAC"
                    SelectedCollection = My.Settings.GACEntries
                    If CurProdGroup.Contains("Data Translation") _
                        Then SelectedCollection = My.Settings.DTGACEntries
                Case "Drivers"
                    SelectedCollection = My.Settings.MCCDriverCategory
                    If CurProdGroup.Contains("Data Translation") _
                        Then SelectedCollection = My.Settings.DTDriverCategory
                Case "Shares"
                    SelectedCollection = My.Settings.SharedFiles
                    If CurProdGroup.Contains("Data Translation") _
                        Then SelectedCollection = My.Settings.DTSharedFiles
            End Select
        End If
        Return SelectedCollection

    End Function

End Class
