Imports System.IO
Imports System.Media
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' OP2OpmTools
' https://github.com/leviathan400/OP2OpmTools
'
' Open OP2MissionEditor .opm json file and
' export as C++ code, json or txt report
'
'
' Outpost 2: Divided Destiny is a real-time strategy video game released in 1997.

Public Class fMain

#Region "Fields and Properties"

    ' Application info
    Public Const ApplicationBuild As String = "0017"
    Public Const ApplicationVersion As String = "0.3.1"
    Public Const ApplicationName As String = "OP2OpmTools"
    Public Const ApplicationDescription As String = "Convert OP2MissionEditor .opm json files to C++ code."
    Public Const ApplicationWebsite As String = "https://github.com/leviathan400/OP2OpmTools"

    Private CurrentOpmData As JObject
    Private MissionDetails As LevelDetails
    Private Units As List(Of Unit)
    Private WallTubes As List(Of WallTube)
    Private Beacons As List(Of Beacon)

#End Region

#Region "Form Events"

    ''' <summary>
    ''' Handles the form load event, initializes the application interface
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">Event arguments</param>
    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Debug.WriteLine("--- " & ApplicationName & " Started ---")
        Me.Icon = My.Resources.Science
        txtConsole.ReadOnly = True
        txtConsole.BackColor = Color.White

        btnExportCpp.Enabled = False
        btnExportJson.Enabled = False
        btnExportTxt.Enabled = False
        btnExportLua.Enabled = False

        Dim player As New SoundPlayer(My.Resources.lab_1)
        player.Play()
    End Sub

    ''' <summary>
    ''' Handles the open button click event
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">Event arguments</param>
    Private Sub txtOpen_Click(sender As Object, e As EventArgs) Handles txtOpen.Click
        OpenOPM()
    End Sub

    ''' <summary>
    ''' Handles the export to C++ button click event
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">Event arguments</param>
    Private Sub btnExportCpp_Click(sender As Object, e As EventArgs) Handles btnExportCpp.Click
        ExportToCpp()
    End Sub

    ''' <summary>
    ''' Handles the export to JSON button click event
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">Event arguments</param>
    Private Sub btnExportJson_Click(sender As Object, e As EventArgs) Handles btnExportJson.Click
        ExportToJson()
    End Sub

    ''' <summary>
    ''' Handles the export to text button click event
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">Event arguments</param>
    Private Sub btnExportTxt_Click(sender As Object, e As EventArgs) Handles btnExportTxt.Click
        ExportToTxt()
    End Sub

    ''' <summary>
    ''' Handles the export to Lua placement button click event
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">Event arguments</param>
    Private Sub btnExportLua_Click(sender As Object, e As EventArgs) Handles btnExportLua.Click
        ExportToLua()
    End Sub

#End Region

#Region "Utility Methods"

    ''' <summary>
    ''' Appends text to the console TextBox with a newline and scrolls to the latest text
    ''' </summary>
    ''' <param name="text">The text to append to the console</param>
    Private Sub AppendToConsole(text As String)
        ' Append text to the console with a newline
        txtConsole.AppendText(text & vbCrLf)
        ' Scroll to the end of the text
        txtConsole.SelectionStart = txtConsole.Text.Length
        txtConsole.ScrollToCaret()
    End Sub

#End Region

#Region "File Operations"

    ''' <summary>
    ''' Opens and parses an Outpost 2 mission (.opm) file
    ''' </summary>
    Private Sub OpenOPM()
        ' Create OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Outpost 2 Mission Files (*.opm)|*.opm"
        openFileDialog.Title = "Open Outpost 2 Mission File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' Clear the console
                txtConsole.Clear()

                ' Read the JSON file
                Dim jsonContent As String = File.ReadAllText(openFileDialog.FileName)

                ' Parse the JSON content
                CurrentOpmData = JObject.Parse(jsonContent)

                ' Process each section
                ProcessLevelDetails()
                ProcessUnits()
                ProcessWallTubes()
                ProcessBeacons()

                ' Display summary information
                AppendToConsole("")
                AppendToConsole("File opened: " & Path.GetFileName(openFileDialog.FileName))
                AppendToConsole($"Total units found: {Units.Count}")
                AppendToConsole($"Total wall/tube elements found: {WallTubes.Count}")
                AppendToConsole($"Total beacons found: {Beacons.Count}")
                Debug.WriteLine("File opened: " & Path.GetFileName(openFileDialog.FileName))

                ' Enable the export buttons
                btnExportCpp.Enabled = True
                btnExportJson.Enabled = True
                btnExportTxt.Enabled = True
                btnExportLua.Enabled = True

            Catch ex As Exception
                AppendToConsole("Error opening file: " & ex.Message)
                AppendToConsole(ex.StackTrace)
            End Try
        End If
    End Sub

#End Region

#Region "Data Processing"

    ''' <summary>
    ''' Processes the level details section from the OPM data
    ''' </summary>
    Private Sub ProcessLevelDetails()
        ' Process LevelDetails
        Dim levelDetailsJson As JObject = TryCast(CurrentOpmData("LevelDetails"), JObject)

        If levelDetailsJson IsNot Nothing Then
            MissionDetails = JsonConvert.DeserializeObject(Of LevelDetails)(levelDetailsJson.ToString())

            AppendToConsole("Mission: " & MissionDetails.ToString())
            AppendToConsole("Description: " & If(String.IsNullOrEmpty(MissionDetails.LevelDescription), "None", MissionDetails.LevelDescription))
            AppendToConsole("Tech Level: " & MissionDetails.MaxTechLevel)
            AppendToConsole("Tech Tree: " & MissionDetails.TechTreeName)
            AppendToConsole("Unit Only: " & MissionDetails.UnitOnlyMission)
            AppendToConsole("")
        Else
            MissionDetails = Nothing
            AppendToConsole("No level details found")
        End If
    End Sub

    ''' <summary>
    ''' Processes the units section from the OPM data for all players
    ''' </summary>
    Private Sub ProcessUnits()
        ' Initialize the units list
        Units = New List(Of Unit)

        ' Get the players array from the MasterVariant section
        Dim players As JArray = CurrentOpmData("MasterVariant")("Players")

        ' Loop through each player
        Dim playerIndex As Integer = 0
        For Each player As JObject In players
            ' Get the player color
            Dim playerColor As String = player("Color").ToString()

            ' Get the units array for the current player
            Dim playerUnits As JArray = TryCast(player("Resources")("Units"), JArray)

            ' Process units if any
            If playerUnits IsNot Nothing AndAlso playerUnits.Count > 0 Then
                AppendToConsole($"Player {playerIndex} ({playerColor}) - {playerUnits.Count} units")

                For Each unitJson As JObject In playerUnits
                    ' Deserialize the unit JSON to our Unit class
                    Dim unit As Unit = JsonConvert.DeserializeObject(Of Unit)(unitJson.ToString())

                    ' Add player information to unit
                    unit.PlayerIndex = playerIndex
                    unit.PlayerColor = playerColor

                    ' Add the unit to our list
                    Units.Add(unit)
                Next
            End If

            playerIndex += 1
        Next

        ' Display unit information in console
        If Units.Count > 0 Then
            AppendToConsole("")
            AppendToConsole("--- Units ---")
            For Each unit As Unit In Units
                Dim details As String = $"Player {unit.PlayerIndex} ({unit.PlayerColor}): {unit.TypeID} at {unit.Position.X}, {unit.Position.Y}"

                ' Add direction if present
                If Not String.IsNullOrEmpty(unit.Direction) Then
                    details &= $", Direction: {unit.Direction}"
                End If

                ' Add cargo info if present
                If unit.CargoType IsNot Nothing AndAlso unit.CargoType <> "None" AndAlso unit.CargoType <> "Empty" Then
                    details &= $", Cargo: {unit.CargoType}"

                    If unit.CargoAmount > 0 Then
                        details &= $" (Amount: {unit.CargoAmount})"
                    End If
                End If

                ' Add lights info if on
                If unit.Lights Then
                    details &= ", Lights: On"
                End If

                ' Add health if not default
                If unit.Health <> 1 Then
                    details &= $", Health: {unit.Health}"
                End If

                AppendToConsole(details)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Processes the wall/tube elements from the OPM data for all players
    ''' </summary>
    Private Sub ProcessWallTubes()
        ' Initialize the wall tubes list
        WallTubes = New List(Of WallTube)

        ' Get the players array from the MasterVariant section
        Dim players As JArray = CurrentOpmData("MasterVariant")("Players")

        ' Loop through each player
        Dim playerIndex As Integer = 0
        For Each player As JObject In players
            ' Get the walltubes array for the current player
            Dim playerWallTubes As JArray = TryCast(player("Resources")("WallTubes"), JArray)

            ' Process walltubes if any
            If playerWallTubes IsNot Nothing AndAlso playerWallTubes.Count > 0 Then
                AppendToConsole($"Found {playerWallTubes.Count} wall/tube elements in player {playerIndex}'s section")

                For Each wallTubeJson As JObject In playerWallTubes
                    ' Deserialize the walltube JSON to our WallTube class
                    Dim wallTube As WallTube = JsonConvert.DeserializeObject(Of WallTube)(wallTubeJson.ToString())

                    ' Add the walltube to our list
                    WallTubes.Add(wallTube)
                Next
            End If

            playerIndex += 1
        Next

        ' Display walltube information in console
        If WallTubes.Count > 0 Then
            AppendToConsole("")
            AppendToConsole("--- Wall/Tube Elements ---")
            For Each wallTube As WallTube In WallTubes
                AppendToConsole($"{wallTube.TypeID} at {wallTube.Position.X}, {wallTube.Position.Y}")
            Next
        End If
    End Sub

    ''' <summary>
    ''' Processes the beacon data from the OPM file
    ''' </summary>
    Private Sub ProcessBeacons()
        ' Initialize the beacons list
        Beacons = New List(Of Beacon)

        ' Process Beacons
        Dim beaconsArray As JArray = TryCast(CurrentOpmData("MasterVariant")("TethysGame")("Beacons"), JArray)

        If beaconsArray IsNot Nothing AndAlso beaconsArray.Count > 0 Then
            For Each beaconJson As JObject In beaconsArray
                ' Deserialize the beacon JSON to our Beacon class
                Dim beacon As Beacon = JsonConvert.DeserializeObject(Of Beacon)(beaconJson.ToString())

                ' Add the beacon to our list
                Beacons.Add(beacon)
            Next

            AppendToConsole($"Found {Beacons.Count} beacons")

            ' Display beacon information in console
            AppendToConsole("")
            AppendToConsole("--- Beacons ---")
            For Each beacon As Beacon In Beacons
                AppendToConsole($"{beacon.MapID} ({beacon.OreType}) at {beacon.Position.X}, {beacon.Position.Y}" &
                           $", Variant: {beacon.BarVariant}, Yield: {beacon.BarYield}")
            Next
        Else
            AppendToConsole("No beacons found")
        End If
    End Sub

#End Region

#Region "Export Operations"

    ''' <summary>
    ''' Exports the loaded OPM data as C++ code for Outpost 2 SDK
    ''' </summary>
    Private Sub ExportToCpp()
        If (Units Is Nothing OrElse Units.Count = 0) AndAlso (WallTubes Is Nothing OrElse WallTubes.Count = 0) AndAlso
       (Beacons Is Nothing OrElse Beacons.Count = 0) AndAlso (MissionDetails Is Nothing) Then
            AppendToConsole("No data to export. Please open an OPM file first.")
            Return
        End If

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "C++ Files (*.cpp)|*.cpp"
        saveFileDialog.Title = "Export to C++ Code"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                Using writer As New StreamWriter(saveFileDialog.FileName)
                    ' Write header with metadata
                    writer.WriteLine("// Generated by " & ApplicationName & " v" & ApplicationVersion)
                    writer.WriteLine("// Export date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    writer.WriteLine("// Source: " & If(MissionDetails IsNot Nothing, MissionDetails.MapName, "Unknown"))
                    writer.WriteLine("")

                    ' Write InitProc function beginning
                    writer.WriteLine("Export int InitProc()")
                    writer.WriteLine("{")

                    ' Mission setup if available
                    If MissionDetails IsNot Nothing Then
                        writer.WriteLine($"    // Mission: {MissionDetails.MissionType}, Players: {MissionDetails.NumPlayers}, Tech Level: {MissionDetails.MaxTechLevel}")
                        writer.WriteLine($"    // Map: {MissionDetails.MapName}")
                        writer.WriteLine($"    // Description: {MissionDetails.LevelDescription}")
                        writer.WriteLine("")
                    End If

                    ' Add player setup for each player
                    If Units IsNot Nothing AndAlso Units.Count > 0 Then
                        Dim playerIndices = Units.Select(Function(u) u.PlayerIndex).Distinct().OrderBy(Function(x) x).ToList()

                        For Each playerIdx In playerIndices
                            Dim playerColor = Units.First(Function(u) u.PlayerIndex = playerIdx).PlayerColor
                            writer.WriteLine($"    // Player {playerIdx} ({playerColor})")

                            ' Check if Eden or Plymouth colony
                            Dim edenUnits = Units.Where(Function(u) u.PlayerIndex = playerIdx AndAlso Not String.IsNullOrEmpty(u.TypeID) AndAlso u.TypeID.Contains("Tokamak")).Count()
                            Dim plymouthUnits = Units.Where(Function(u) u.PlayerIndex = playerIdx AndAlso Not String.IsNullOrEmpty(u.TypeID) AndAlso u.TypeID.Contains("Geothermal")).Count()

                            If edenUnits > plymouthUnits Then
                                writer.WriteLine($"    Player[{playerIdx}].GoEden();")
                            Else
                                writer.WriteLine($"    Player[{playerIdx}].GoPlymouth();")
                            End If

                            writer.WriteLine($"    Player[{playerIdx}].SetColorNumber({GetColorNumber(playerColor)});")
                            writer.WriteLine("")
                        Next
                    End If

                    ' Create a variable for unit reference
                    writer.WriteLine("    Unit x;")
                    writer.WriteLine("")

                    ' Units
                    If Units IsNot Nothing AndAlso Units.Count > 0 Then
                        writer.WriteLine("    // Create units")

                        For Each unit As Unit In Units
                            ' Build enhanced comment for unit
                            Dim comment As String = $"    // {unit.TypeID}"

                            ' Add cargo information for cargo vehicles
                            If unit.CargoType IsNot Nothing AndAlso unit.CargoType <> "None" AndAlso unit.CargoType <> "Empty" Then
                                If unit.TypeID = "CargoTruck" Then
                                    ' Check if this is a starship part (either by CargoType being numeric or CargoAmount)
                                    Dim isStarshipPart As Boolean = False
                                    Dim starshipPartId As Integer = 0

                                    If IsNumeric(unit.CargoType) Then
                                        starshipPartId = Integer.Parse(unit.CargoType)
                                        isStarshipPart = (starshipPartId >= 88 AndAlso starshipPartId <= 104)
                                    ElseIf unit.CargoType = "Spaceport" AndAlso unit.CargoAmount >= 88 AndAlso unit.CargoAmount <= 104 Then
                                        starshipPartId = unit.CargoAmount
                                        isStarshipPart = True
                                    End If

                                    If isStarshipPart Then
                                        comment &= $" with {GetStarshipPartName(starshipPartId)}"
                                    Else
                                        comment &= $" with {unit.CargoType}"

                                        ' Add cargo amount if present and not a starship part
                                        If unit.CargoAmount > 0 Then
                                            comment &= $" (amount: {unit.CargoAmount})"
                                        End If
                                    End If
                                ElseIf unit.TypeID = "ConVec" Then
                                    comment &= $" carrying {unit.CargoType}"
                                ElseIf unit.TypeID = "Lynx" OrElse unit.TypeID = "Panther" OrElse unit.TypeID = "Tiger" OrElse
                                    unit.TypeID = "Spider" OrElse unit.TypeID = "Scorpion" OrElse unit.TypeID = "GuardPost" Then
                                    comment &= $" {unit.CargoType}"
                                End If
                            End If

                            ' Add position
                            comment &= $" at position {unit.Position.X}, {unit.Position.Y}"

                            ' Add direction if available
                            If Not String.IsNullOrEmpty(unit.Direction) Then
                                comment &= $", facing {unit.Direction}"
                            End If

                            writer.WriteLine(comment)

                            ' Write the CreateUnit call with transformation formula
                            Dim createCall = $"    TethysGame::CreateUnit(x, map{unit.TypeID}, LOCATION({unit.Position.X} + 31, {unit.Position.Y} - 1), {unit.PlayerIndex}"

                            ' Add cargo/weapon type based on unit type
                            If unit.CargoType IsNot Nothing AndAlso unit.CargoType <> "None" AndAlso unit.CargoType <> "Empty" Then
                                ' For combat units and ConVecs, include cargo/weapon in CreateUnit call
                                If unit.TypeID = "Lynx" OrElse unit.TypeID = "Panther" OrElse unit.TypeID = "Tiger" OrElse
               unit.TypeID = "Spider" OrElse unit.TypeID = "Scorpion" OrElse unit.TypeID = "GuardPost" OrElse
               unit.TypeID = "ConVec" Then
                                    createCall &= $", map{unit.CargoType}"
                                Else
                                    createCall &= ", mapNone"
                                End If
                            Else
                                createCall &= ", mapNone"
                            End If

                            ' Add direction/rotation
                            If Not String.IsNullOrEmpty(unit.Direction) Then
                                createCall &= $", {GetDirectionValue(unit.Direction)}"
                            Else
                                createCall &= ", 0" ' Default direction (East)
                            End If

                            createCall &= ");"
                            writer.WriteLine(createCall)

                            ' Add lights setting if needed
                            If unit.Lights Then
                                writer.WriteLine("    x.DoSetLights(1);")
                            End If

                            ' Handle special cargo types
                            If unit.CargoType IsNot Nothing AndAlso unit.CargoType <> "None" AndAlso unit.CargoType <> "Empty" Then
                                ' Add cargo for CargoTruck
                                If unit.TypeID = "CargoTruck" Then
                                    Dim truckCargoType = GetTruckCargoType(unit.CargoType)

                                    ' Handle different cargo types
                                    If truckCargoType = "truckSpaceport" Then
                                        ' For spaceport modules, when cargo is a number, it's likely the module ID
                                        Dim moduleId As Integer = 0

                                        If IsNumeric(unit.CargoType) Then
                                            moduleId = Integer.Parse(unit.CargoType)
                                        ElseIf unit.CargoAmount > 0 Then
                                            moduleId = unit.CargoAmount  ' Module ID stored in CargoAmount
                                        Else
                                            moduleId = 90  ' Default to IonDriveModule if no ID specified
                                        End If

                                        writer.WriteLine($"    x.SetTruckCargo({truckCargoType}, {moduleId});")
                                    Else
                                        ' For normal cargo, use 1000 as default amount if not specified
                                        Dim cargoAmount = If(unit.CargoAmount > 0, unit.CargoAmount, 1000)
                                        writer.WriteLine($"    x.SetTruckCargo({truckCargoType}, {cargoAmount});")
                                    End If
                                    ' Handle StructureFactory cargo (NOT ConVec)
                                ElseIf unit.TypeID = "StructureFactory" OrElse unit.TypeID = "VehicleFactory" Then
                                    ' Map the building type to its corresponding enum
                                    If IsValidBuilding(unit.CargoType) Then
                                        Dim cargoAmount = If(unit.CargoAmount > 0, unit.CargoAmount, 0)
                                        writer.WriteLine($"    x.SetFactoryCargo(0, map{unit.CargoType}, mapNone);")
                                    End If
                                End If
                            End If
                        Next

                        writer.WriteLine("")
                    End If

                    ' Wall/Tubes
                    If WallTubes IsNot Nothing AndAlso WallTubes.Count > 0 Then
                        writer.WriteLine("    // Create walls and tubes")

                        ' Group wall/tubes by type and adjacent positions for CreateTubeOrWallLine calls
                        Dim wallGroups = GroupWallTubesIntoLines(WallTubes)

                        For Each group In wallGroups
                            writer.WriteLine($"    CreateTubeOrWallLine({group.StartX} + 31, {group.StartY} - 1, {group.EndX} + 31, {group.EndY} - 1, map{group.Type});")
                        Next

                        ' Individual wall/tubes that couldn't be grouped
                        Dim individualWallTubes = GetIndividualWallTubes(WallTubes, wallGroups)

                        For Each wallTube In individualWallTubes
                            writer.WriteLine($"    TethysGame::CreateWallOrTube({wallTube.Position.X} + 31, {wallTube.Position.Y} - 1, 0, map{wallTube.TypeID});")
                        Next

                        writer.WriteLine("")
                    End If

                    ' Beacons
                    If Beacons IsNot Nothing AndAlso Beacons.Count > 0 Then
                        writer.WriteLine("    // Create beacons")

                        For Each beacon In Beacons
                            ' Convert MapID to proper type (MiningBeacon, MagmaVent, Fumarole)
                            Dim beaconType = beacon.MapID

                            ' OreType: Common = 0, Rare = 1, Random = -1
                            Dim oreTypeValue As Integer
                            Select Case beacon.OreType.ToLower()
                                Case "common"
                                    oreTypeValue = 0 ' OreTypeCommon
                                Case "rare"
                                    oreTypeValue = 1 ' OreTypeRare
                                Case "random"
                                    oreTypeValue = -1 ' OreTypeRandom
                                Case Else
                                    oreTypeValue = 0 ' Default to common
                            End Select

                            ' BarYield: Bar3 = 0, Bar2 = 1, Bar1 = 2, BarRandom = -1
                            Dim yieldValue As Integer = GetYieldValue(beacon.BarYield)

                            ' BarVariant: Variant1 = 0, Variant2 = 1, Variant3 = 2, VariantRandom = -1
                            Dim variantValue As Integer = GetVariantValue(beacon.BarVariant)

                            writer.WriteLine($"    TethysGame::CreateBeacon(map{beaconType}, {beacon.Position.X} + 31, {beacon.Position.Y} - 1, {oreTypeValue}, {yieldValue}, {variantValue});")
                        Next

                        writer.WriteLine("")
                    End If

                    ' Return success
                    writer.WriteLine("    return 1; // return 1 if OK; 0 on failure")
                    writer.WriteLine("}")
                End Using

                AppendToConsole("C++ code exported successfully to: " & saveFileDialog.FileName)
                Process.Start(saveFileDialog.FileName)

            Catch ex As Exception
                AppendToConsole("Error exporting to C++: " & ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Exports the loaded OPM data as a complete OP2Lua mission - BOTH script files that make up a
    ''' mission, written side by side with a shared base name:
    '''   &lt;base&gt;.placement.lua  - the layout table OP2LuaCore consumes (players/units/beacons/walls/...)
    '''   &lt;base&gt;.lua            - a starter mission-logic script (on_init with win/lose conditions)
    ''' Together with a stub DLL (cloned by the OP2LuaEditor / new-mission tool) these are a runnable
    ''' mission - this is the .opm -&gt; Lua converter the SDK docs refer to.
    '''
    ''' COORDINATES: OP2Lua placement coordinates are the SAME pre-offset editor coordinates stored
    ''' in the .opm - the OP2Lua runtime applies the +31 X / -1 Y engine offset itself. So, unlike
    ''' the C++ export (which bakes "+ 31 / - 1" into LOCATION(...)), we write the raw .opm X/Y
    ''' straight through with no transformation.
    ''' </summary>
    Private Sub ExportToLua()
        If (Units Is Nothing OrElse Units.Count = 0) AndAlso (WallTubes Is Nothing OrElse WallTubes.Count = 0) AndAlso
           (Beacons Is Nothing OrElse Beacons.Count = 0) AndAlso (MissionDetails Is Nothing) Then
            AppendToConsole("No data to export. Please open an OPM file first.")
            Return
        End If

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Lua Mission Files (*.lua)|*.lua"
        saveFileDialog.Title = "Export to Lua Mission (creates both .lua + .placement.lua)"
        saveFileDialog.FileName = "mission.lua"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' Derive a shared base name from the chosen file, so both scripts share it (the stub
                ' DLL loads <base>.lua + <base>.placement.lua). Strip a trailing ".placement" so that
                ' picking "MyMission.placement.lua" still yields base "MyMission".
                Dim outDir As String = Path.GetDirectoryName(saveFileDialog.FileName)
                Dim baseName As String = Path.GetFileNameWithoutExtension(saveFileDialog.FileName)
                If baseName.EndsWith(".placement", StringComparison.OrdinalIgnoreCase) Then
                    baseName = baseName.Substring(0, baseName.Length - ".placement".Length)
                End If
                If String.IsNullOrEmpty(baseName) Then baseName = "mission"

                Dim missionName As String = baseName   ' menu/display name (the author edits this later)
                Dim placementPath As String = Path.Combine(outDir, baseName & ".placement.lua")
                Dim missionPath As String = Path.Combine(outDir, baseName & ".lua")

                Dim mapName As String = If(MissionDetails IsNot Nothing, MissionDetails.MapName, "")
                Dim techName As String = If(MissionDetails IsNot Nothing, MissionDetails.TechTreeName, "MULTITEK.TXT")
                Dim missionType As String = If(MissionDetails IsNot Nothing, MissionDetails.MissionType, "Colony")
                Dim maxTech As Integer = If(MissionDetails IsNot Nothing, MissionDetails.MaxTechLevel, 12)
                Dim numPlayers As Integer = If(MissionDetails IsNot Nothing, MissionDetails.NumPlayers, 0)

                ' --- 1) <base>.placement.lua : the layout, straight from the .opm ---
                Using writer As New StreamWriter(placementPath)
                    writer.WriteLine("-- " & baseName & ".placement.lua  -  generated by " & ApplicationName & " v" & ApplicationVersion &
                                     " from " & If(String.IsNullOrEmpty(mapName), "an .opm file", mapName))
                    writer.WriteLine("-- Coordinates are the in-game status-bar coords stored in the .opm (no offset applied).")
                    writer.WriteLine("return {")
                    writer.WriteLine($"  name = {LuaStr(missionName)}, map = {LuaStr(mapName)}, tech = {LuaStr(techName)}, type = {LuaStr(missionType)},")
                    writer.WriteLine($"  max_tech = {maxTech}, players_count = {numPlayers},")
                    writer.WriteLine("")

                    WriteLuaPlayers(writer)
                    WriteLuaUnits(writer)
                    WriteLuaBeacons(writer)
                    WriteLuaWalls(writer)

                    ' Regions aren't stored in the .opm (OP2MissionEditor has no region authoring UI),
                    ' so emit an empty table for the author to fill in.
                    writer.WriteLine("  -- regions are not stored in the .opm; add named rectangles { x1, y1, x2, y2 }")
                    writer.WriteLine("  -- or points { x, y } that your mission script references.")
                    writer.WriteLine("  regions = {},")

                    WriteLuaMarkers(writer)

                    writer.WriteLine("}")
                End Using

                ' --- 2) <base>.lua : a starter mission-logic script ---
                Using writer As New StreamWriter(missionPath)
                    WriteLuaMission(writer, baseName, missionName, mapName)
                End Using

                AppendToConsole("Lua mission exported (2 files):")
                AppendToConsole("  " & missionPath & "    (mission logic - edit this)")
                AppendToConsole("  " & placementPath & "    (layout from the .opm)")
                AppendToConsole("Drop both next to a stub DLL named " & baseName & ".dll to run the mission.")
                Process.Start("explorer.exe", "/select,""" & missionPath & """")

            Catch ex As Exception
                AppendToConsole("Error exporting to Lua: " & ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Writes the players table for the Lua placement, read directly from the OPM JSON so we pick up
    ''' colony (IsEden), human (IsHuman), color, bot type and starting resources - data the shared
    ''' Unit list doesn't carry. Player index is written 1-based (OP2Lua players are [1], [2], ...).
    ''' </summary>
    Private Sub WriteLuaPlayers(writer As StreamWriter)
        Dim master As JObject = TryCast(CurrentOpmData("MasterVariant"), JObject)
        Dim players As JArray = If(master Is Nothing, Nothing, TryCast(master("Players"), JArray))

        If players Is Nothing OrElse players.Count = 0 Then
            writer.WriteLine("  players = {},")
            writer.WriteLine("")
            Return
        End If

        writer.WriteLine("  players = {")

        Dim fallbackIndex As Integer = 0
        For Each player As JObject In players
            Dim luaIndex As Integer = JInt(player, "ID", fallbackIndex) + 1
            Dim colony As String = If(JBool(player, "IsEden", True), "Eden", "Plymouth")
            Dim human As Boolean = JBool(player, "IsHuman", False)
            Dim color As String = JStr(player, "Color", "Blue")
            Dim botType As String = JStr(player, "BotType", "")

            writer.WriteLine($"    [{luaIndex}] = {{")

            Dim header As String = $"      colony = {LuaStr(colony)}, human = {LuaBool(human)}, color = {LuaStr(color)}"
            If botType <> "" AndAlso botType <> "None" Then
                header &= $", bot = {LuaStr(botType)}"
            End If
            writer.WriteLine(header & ",")

            Dim res As JObject = TryCast(player("Resources"), JObject)
            If res IsNot Nothing Then
                Dim parts As New List(Of String) From {
                    $"common_ore = {JInt(res, "CommonOre", 0)}",
                    $"rare_ore = {JInt(res, "RareOre", 0)}",
                    $"food = {JInt(res, "Food", 0)}",
                    $"kids = {JInt(res, "Kids", 0)}",
                    $"workers = {JInt(res, "Workers", 0)}",
                    $"scientists = {JInt(res, "Scientists", 0)}",
                    $"tech_level = {JInt(res, "TechLevel", 0)}"
                }
                Dim morale As String = JStr(res, "MoraleLevel", "")
                If morale <> "" Then parts.Add($"morale = {LuaStr(morale)}")
                writer.WriteLine("      resources = { " & String.Join(", ", parts) & " },")

                Dim centerView As JObject = TryCast(res("CenterView"), JObject)
                If centerView IsNot Nothing Then
                    Dim cvx As Integer = JInt(centerView, "X", 0)
                    Dim cvy As Integer = JInt(centerView, "Y", 0)
                    If cvx <> 0 OrElse cvy <> 0 Then
                        writer.WriteLine($"      center_view = {{ {cvx}, {cvy} }},")
                    End If
                End If
            End If

            writer.WriteLine("    },")
            fallbackIndex += 1
        Next

        writer.WriteLine("  },")
        writer.WriteLine("")
    End Sub

    ''' <summary>
    ''' Writes the units table. Player is 1-based; combat units emit their weapon, cargo carriers
    ''' emit cargo (+ amount). Coordinates are passed through unchanged (see ExportToLua remarks).
    ''' </summary>
    Private Sub WriteLuaUnits(writer As StreamWriter)
        If Units Is Nothing OrElse Units.Count = 0 Then
            writer.WriteLine("  units = {},")
            writer.WriteLine("")
            Return
        End If

        writer.WriteLine("  units = {")
        For Each unit As Unit In Units
            Dim parts As New List(Of String) From {
                $"type = {LuaStr(unit.TypeID)}",
                $"player = {unit.PlayerIndex + 1}",
                $"at = {{ {unit.Position.X}, {unit.Position.Y} }}"
            }

            Dim hasCargo As Boolean = unit.CargoType IsNot Nothing AndAlso unit.CargoType <> "" AndAlso
                                      unit.CargoType <> "None" AndAlso unit.CargoType <> "Empty"
            If hasCargo Then
                If IsCombatUnit(unit.TypeID) Then
                    parts.Add($"weapon = {LuaStr(unit.CargoType)}")
                ElseIf unit.TypeID = "CargoTruck" OrElse unit.TypeID = "ConVec" Then
                    parts.Add($"cargo = {LuaStr(unit.CargoType)}")
                    If unit.CargoAmount > 0 Then parts.Add($"amount = {unit.CargoAmount}")
                End If
            End If

            writer.WriteLine("    { " & String.Join(", ", parts) & " },")
        Next
        writer.WriteLine("  },")
        writer.WriteLine("")
    End Sub

    ''' <summary>
    ''' Writes the beacons table ({ type, ore (mining only), at }).
    ''' </summary>
    Private Sub WriteLuaBeacons(writer As StreamWriter)
        If Beacons Is Nothing OrElse Beacons.Count = 0 Then
            writer.WriteLine("  beacons = {},")
            writer.WriteLine("")
            Return
        End If

        writer.WriteLine("  beacons = {")
        For Each beacon As Beacon In Beacons
            Dim parts As New List(Of String) From {$"type = {LuaStr(beacon.MapID)}"}
            If beacon.MapID = "MiningBeacon" AndAlso Not String.IsNullOrEmpty(beacon.OreType) Then
                parts.Add($"ore = {LuaStr(beacon.OreType)}")
            End If
            parts.Add($"at = {{ {beacon.Position.X}, {beacon.Position.Y} }}")
            writer.WriteLine("    { " & String.Join(", ", parts) & " },")
        Next
        writer.WriteLine("  },")
        writer.WriteLine("")
    End Sub

    ''' <summary>
    ''' Writes the walls table - one entry per wall/tube tile ({ type, at }).
    ''' </summary>
    Private Sub WriteLuaWalls(writer As StreamWriter)
        If WallTubes Is Nothing OrElse WallTubes.Count = 0 Then
            writer.WriteLine("  walls = {},")
            writer.WriteLine("")
            Return
        End If

        writer.WriteLine("  walls = {")
        For Each wallTube As WallTube In WallTubes
            writer.WriteLine($"    {{ type = {LuaStr(wallTube.TypeID)}, at = {{ {wallTube.Position.X}, {wallTube.Position.Y} }} }},")
        Next
        writer.WriteLine("  },")
        writer.WriteLine("")
    End Sub

    ''' <summary>
    ''' Writes the markers table, read straight from MasterVariant.TethysGame.Markers. Markers have no
    ''' name in the .opm, so they're emitted as marker_1, marker_2, ... (the MarkerType is a comment).
    ''' </summary>
    Private Sub WriteLuaMarkers(writer As StreamWriter)
        Dim master As JObject = TryCast(CurrentOpmData("MasterVariant"), JObject)
        Dim tethysGame As JObject = If(master Is Nothing, Nothing, TryCast(master("TethysGame"), JObject))
        Dim markers As JArray = If(tethysGame Is Nothing, Nothing, TryCast(tethysGame("Markers"), JArray))

        If markers Is Nothing OrElse markers.Count = 0 Then
            writer.WriteLine("  markers = {},")
            Return
        End If

        writer.WriteLine("  markers = {")
        Dim index As Integer = 1
        For Each marker As JObject In markers
            Dim pos As JObject = TryCast(marker("Position"), JObject)
            Dim x As Integer = JInt(pos, "X", 0)
            Dim y As Integer = JInt(pos, "Y", 0)
            Dim markerType As String = JStr(marker, "MarkerType", "")
            writer.WriteLine($"    marker_{index} = {{ at = {{ {x}, {y} }} }},  -- {markerType}")
            index += 1
        Next
        writer.WriteLine("  },")
    End Sub

    ''' <summary>
    ''' Writes a starter &lt;base&gt;.lua mission-logic script. The .opm carries no mission logic, so this
    ''' is a sensible skeleton: an on_init with an objective message and win/lose conditions wired to
    ''' the actual human/enemy players found in the data (destroy the enemy Command Center / don't lose
    ''' your own). The author fleshes out the rest using docs/API.md.
    ''' </summary>
    Private Sub WriteLuaMission(writer As StreamWriter, baseName As String, missionName As String, mapName As String)
        ' Find the first human and first non-human (enemy) player, 1-based, from the OPM.
        Dim humanIndex As Integer = 0
        Dim enemyIndex As Integer = 0
        Dim master As JObject = TryCast(CurrentOpmData("MasterVariant"), JObject)
        Dim players As JArray = If(master Is Nothing, Nothing, TryCast(master("Players"), JArray))
        Dim fallbackIndex As Integer = 0
        If players IsNot Nothing Then
            For Each player As JObject In players
                Dim idx As Integer = JInt(player, "ID", fallbackIndex) + 1
                If JBool(player, "IsHuman", False) Then
                    If humanIndex = 0 Then humanIndex = idx
                Else
                    If enemyIndex = 0 Then enemyIndex = idx
                End If
                fallbackIndex += 1
            Next
        End If
        If humanIndex = 0 Then humanIndex = 1

        ' Only reference a player:owns_any("CommandCenter") condition when that player actually has one,
        ' so the generated mission doesn't fire on an impossible/empty condition.
        Dim enemyHasCC As Boolean = enemyIndex > 0 AndAlso Units IsNot Nothing AndAlso
            Units.Any(Function(u) u.PlayerIndex + 1 = enemyIndex AndAlso u.TypeID = "CommandCenter")
        Dim humanHasCC As Boolean = Units IsNot Nothing AndAlso
            Units.Any(Function(u) u.PlayerIndex + 1 = humanIndex AndAlso u.TypeID = "CommandCenter")

        writer.WriteLine("-- " & baseName & ".lua  -  " & missionName & "  -  mission logic.")
        writer.WriteLine("-- Generated by " & ApplicationName & " v" & ApplicationVersion & " from " &
                         If(String.IsNullOrEmpty(mapName), "an .opm file", mapName) & ".")
        writer.WriteLine("-- The .opm has no mission logic, so this is a starter script - edit freely. See docs/API.md.")
        writer.WriteLine("")
        writer.WriteLine("function on_init()")
        writer.WriteLine("  game.message(" & LuaStr(missionName & ": destroy the enemy Command Center!") & ")")
        writer.WriteLine("  game.morale_steady()   -- keep morale stable so it's about the fight, not colony micro")
        writer.WriteLine("")

        If enemyHasCC Then
            writer.WriteLine("  -- WIN: the enemy has no Command Center left.")
            writer.WriteLine($"  when(function() return not players[{enemyIndex}]:owns_any(""CommandCenter"") end, function()")
            writer.WriteLine("    game.message(" & LuaStr("Enemy base destroyed - mission accomplished!") & ")")
            writer.WriteLine("    mission.win()")
            writer.WriteLine("  end)")
        Else
            writer.WriteLine("  -- WIN: define your victory condition, e.g.:")
            writer.WriteLine("  -- when(function() return <condition> end, function() mission.win() end)")
        End If
        writer.WriteLine("")

        If humanHasCC Then
            writer.WriteLine("  -- LOSE: your Command Center is gone.")
            writer.WriteLine($"  when(function() return not players[{humanIndex}]:owns_any(""CommandCenter"") end, function()")
            writer.WriteLine("    game.message(" & LuaStr("Your Command Center is destroyed. Mission failed.") & ")")
            writer.WriteLine("    mission.lose()")
            writer.WriteLine("  end)")
        Else
            writer.WriteLine("  -- LOSE: define your failure condition, e.g.:")
            writer.WriteLine("  -- when(function() return <condition> end, function() mission.lose() end)")
        End If

        writer.WriteLine("end")
    End Sub

    ' --- Lua export helpers ---

    ''' <summary>
    ''' Quotes and escapes a string as a Lua string literal (handles backslash and double-quote).
    ''' </summary>
    Private Function LuaStr(value As String) As String
        If value Is Nothing Then Return "nil"
        Return """" & value.Replace("\", "\\").Replace("""", "\""") & """"
    End Function

    ''' <summary>
    ''' Renders a Boolean as a Lua keyword (true/false).
    ''' </summary>
    Private Function LuaBool(value As Boolean) As String
        Return If(value, "true", "false")
    End Function

    ''' <summary>
    ''' True for unit types that carry a weapon via CargoType (combat chassis + GuardPost).
    ''' </summary>
    Private Function IsCombatUnit(typeId As String) As Boolean
        Select Case typeId
            Case "Lynx", "Panther", "Tiger", "Spider", "Scorpion", "GuardPost"
                Return True
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Reads an integer member from a JObject, tolerating absence/null and float-encoded ints
    ''' (e.g. "1.0"); returns defaultValue when missing or unparseable.
    ''' </summary>
    Private Function JInt(obj As JObject, key As String, defaultValue As Integer) As Integer
        If obj Is Nothing Then Return defaultValue
        Dim token As JToken = obj(key)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return defaultValue
        Try
            Return CInt(token.Value(Of Double)())
        Catch
            Return defaultValue
        End Try
    End Function

    ''' <summary>
    ''' Reads a string member from a JObject, returning defaultValue when missing or null.
    ''' </summary>
    Private Function JStr(obj As JObject, key As String, defaultValue As String) As String
        If obj Is Nothing Then Return defaultValue
        Dim token As JToken = obj(key)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return defaultValue
        Return token.ToString()
    End Function

    ''' <summary>
    ''' Reads a Boolean member from a JObject, returning defaultValue when missing/null/unparseable.
    ''' </summary>
    Private Function JBool(obj As JObject, key As String, defaultValue As Boolean) As Boolean
        If obj Is Nothing Then Return defaultValue
        Dim token As JToken = obj(key)
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return defaultValue
        Try
            Return CBool(token.Value(Of Boolean)())
        Catch
            Return defaultValue
        End Try
    End Function

    ''' <summary>
    ''' Exports the loaded OPM data as JSON format
    ''' </summary>
    Private Sub ExportToJson()
        ' Export to yet another file format?
    End Sub

    ''' <summary>
    ''' Exports the loaded OPM data as a comprehensive text report
    ''' </summary>
    Private Sub ExportToTxt()
        ' Create a report of the OPM
        If (Units Is Nothing OrElse Units.Count = 0) AndAlso (WallTubes Is Nothing OrElse WallTubes.Count = 0) AndAlso
           (Beacons Is Nothing OrElse Beacons.Count = 0) AndAlso (MissionDetails Is Nothing) Then
            AppendToConsole("No data to export. Please open an OPM file first.")
            Return
        End If

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Text Files (*.txt)|*.txt"
        saveFileDialog.Title = "Export to Text Report"

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            Try
                Using writer As New StreamWriter(saveFileDialog.FileName)
                    ' Write header with metadata
                    writer.WriteLine("Outpost 2 Mission Report")
                    writer.WriteLine("Generated by " & ApplicationName & " v" & ApplicationVersion)
                    writer.WriteLine("Date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                    writer.WriteLine(New String("="c, 60))
                    writer.WriteLine()

                    ' Level details section
                    writer.WriteLine("LEVEL DETAILS")
                    writer.WriteLine(New String("-"c, 60))

                    If MissionDetails IsNot Nothing Then
                        writer.WriteLine("Map Name: " & MissionDetails.MapName)
                        writer.WriteLine("Mission Type: " & MissionDetails.MissionType)
                        writer.WriteLine("Description: " & If(String.IsNullOrEmpty(MissionDetails.LevelDescription), "None", MissionDetails.LevelDescription))
                        writer.WriteLine("Number of Players: " & MissionDetails.NumPlayers.ToString())
                        writer.WriteLine("Tech Tree: " & MissionDetails.TechTreeName)
                        writer.WriteLine("Max Tech Level: " & MissionDetails.MaxTechLevel.ToString())
                        writer.WriteLine("Unit Only Mission: " & MissionDetails.UnitOnlyMission.ToString())
                    Else
                        writer.WriteLine("No level details available")
                    End If

                    writer.WriteLine()

                    ' Summary statistics
                    writer.WriteLine("SUMMARY STATISTICS")
                    writer.WriteLine(New String("-"c, 60))

                    ' Player counts
                    Dim playerCount As Integer = 0
                    If Units IsNot Nothing AndAlso Units.Count > 0 Then
                        playerCount = Units.Select(Function(u) u.PlayerIndex).Distinct().Count()
                    End If

                    writer.WriteLine("Total Players: " & playerCount.ToString())
                    writer.WriteLine("Total Units: " & If(Units IsNot Nothing, Units.Count, 0).ToString())
                    writer.WriteLine("Total Wall/Tube Elements: " & If(WallTubes IsNot Nothing, WallTubes.Count, 0).ToString())
                    writer.WriteLine("Total Beacons: " & If(Beacons IsNot Nothing, Beacons.Count, 0).ToString())

                    ' Calculate counts by type
                    If Units IsNot Nothing AndAlso Units.Count > 0 Then
                        Dim unitsByType = Units.GroupBy(Function(u) u.TypeID).
                                        Select(Function(g) New With {.Type = g.Key, .Count = g.Count()}).
                                        OrderByDescending(Function(x) x.Count)

                        writer.WriteLine()
                        writer.WriteLine("Units by Type:")
                        For Each unitType In unitsByType
                            writer.WriteLine("  " & unitType.Type & ": " & unitType.Count.ToString())
                        Next
                    End If

                    ' Wall/Tube elements
                    If WallTubes IsNot Nothing AndAlso WallTubes.Count > 0 Then
                        writer.WriteLine()
                        writer.WriteLine("WALL/TUBE ELEMENTS")
                        writer.WriteLine(New String("-"c, 60))

                        ' Count by type
                        Dim wallTubesByType = WallTubes.GroupBy(Function(w) w.TypeID).
                                             Select(Function(g) New With {.Type = g.Key, .Count = g.Count()}).
                                             OrderByDescending(Function(x) x.Count)

                        writer.WriteLine("Total Wall/Tube Elements: " & WallTubes.Count.ToString())
                        writer.WriteLine()
                        writer.WriteLine("Counts by Type:")

                        For Each wallType In wallTubesByType
                            writer.WriteLine("  " & wallType.Type & ": " & wallType.Count.ToString())
                        Next
                    End If

                    If Beacons IsNot Nothing AndAlso Beacons.Count > 0 Then
                        writer.WriteLine()
                        writer.WriteLine("BEACONS")
                        writer.WriteLine(New String("-"c, 60))

                        ' Count by type
                        Dim beaconsByType = Beacons.GroupBy(Function(b) b.MapID).
                                           Select(Function(g) New With {.Type = g.Key, .Count = g.Count()}).
                                           OrderByDescending(Function(x) x.Count)

                        writer.WriteLine("Total Beacons: " & Beacons.Count.ToString())
                        writer.WriteLine()
                        writer.WriteLine("Counts by Type:")

                        For Each beaconType In beaconsByType
                            writer.WriteLine("  " & beaconType.Type & ": " & beaconType.Count.ToString())
                        Next

                        ' For MiningBeacons, count ore types
                        Dim miningBeacons = Beacons.Where(Function(b) b.MapID = "MiningBeacon").ToList()
                        If miningBeacons.Count > 0 Then
                            writer.WriteLine()
                            writer.WriteLine("Mining Beacons by Ore Type:")
                            Dim beaconsByOreType = miningBeacons.GroupBy(Function(b) b.OreType).
                                                 Select(Function(g) New With {.Type = g.Key, .Count = g.Count()}).
                                                 OrderByDescending(Function(x) x.Count)

                            For Each oreType In beaconsByOreType
                                writer.WriteLine("  " & oreType.Type & ": " & oreType.Count.ToString())
                            Next
                        End If
                    End If

                    ' Detailed player sections
                    If Units IsNot Nothing AndAlso Units.Count > 0 Then
                        writer.WriteLine()
                        writer.WriteLine("PLAYER DETAILS")
                        writer.WriteLine(New String("-"c, 60))

                        Dim playerIndices = Units.Select(Function(u) u.PlayerIndex).Distinct().OrderBy(Function(x) x).ToList()

                        For Each playerIdx In playerIndices
                            Dim playerColor = Units.First(Function(u) u.PlayerIndex = playerIdx).PlayerColor
                            writer.WriteLine("Player " & playerIdx.ToString() & " (" & playerColor & ")")

                            ' Count units for this player
                            Dim playerUnits = Units.Where(Function(u) u.PlayerIndex = playerIdx).ToList()
                            writer.WriteLine("  Total Units: " & playerUnits.Count.ToString())

                            ' Count unit types for this player
                            Dim playerUnitsByType = playerUnits.GroupBy(Function(u) u.TypeID).
                                                  Select(Function(g) New With {.Type = g.Key, .Count = g.Count()}).
                                                  OrderByDescending(Function(x) x.Count)

                            writer.WriteLine("  Units by Type:")
                            For Each unitType In playerUnitsByType
                                writer.WriteLine("    " & unitType.Type & ": " & unitType.Count.ToString())
                            Next

                            writer.WriteLine()
                        Next
                    End If

                    writer.WriteLine("END OF REPORT")
                End Using

                AppendToConsole("Text report exported successfully to: " & saveFileDialog.FileName)
                Process.Start(saveFileDialog.FileName)

            Catch ex As Exception
                AppendToConsole("Error exporting to text: " & ex.Message)
            End Try
        End If
    End Sub

#End Region

#Region "Helper Functions - Validation"

    ''' <summary>
    ''' Checks if a string represents a valid building type for Outpost 2
    ''' </summary>
    ''' <param name="buildingTypeStr">The building type string to validate</param>
    ''' <returns>True if the building type is valid, False otherwise</returns>
    Private Function IsValidBuilding(buildingTypeStr As String) As Boolean
        ' List of known valid building types
        Dim validBuildings As String() = {
        "CommonOreMine", "RareOreMine", "GuardPost", "LightTower", "CommonStorage", "RareStorage",
        "Forum", "CommandCenter", "MHDGenerator", "Residence", "RobotCommand", "TradeCenter",
        "BasicLab", "MedicalCenter", "Nursery", "SolarPowerArray", "RecreationFacility", "University",
        "Agridome", "DIRT", "Garage", "MagmaWell", "MeteorDefense", "GeothermalPlant", "ArachnidFactory",
        "ConsumerFactory", "StructureFactory", "VehicleFactory", "StandardLab", "AdvancedLab", "Observatory",
        "ReinforcedResidence", "AdvancedResidence", "CommonOreSmelter", "Spaceport", "RareOreSmelter",
        "GORF", "Tokamak"
    }

        Return Array.IndexOf(validBuildings, buildingTypeStr) >= 0
    End Function

    ''' <summary>
    ''' Checks if a string represents a valid map_id enum value
    ''' </summary>
    ''' <param name="mapIdStr">The map ID string to validate</param>
    ''' <returns>True if the map ID is valid, False otherwise</returns>
    Private Function IsValidMapId(mapIdStr As String) As Boolean
        ' List of known valid map_id values that might be used as cargo
        Dim validMapIds As String() = {"Spaceport", "CommandCenter", "Tokamak", "CommonOreSmelter", "RareOreSmelter",
                                  "Agridome", "StructureFactory", "VehicleFactory", "CommonStorage", "RareStorage"}

        Return Array.IndexOf(validMapIds, mapIdStr) >= 0
    End Function

#End Region

#Region "Helper Functions - Game Data Conversion"

    ''' <summary>
    ''' Converts a color name to its corresponding numeric value for Outpost 2
    ''' </summary>
    ''' <param name="colorName">The color name to convert</param>
    ''' <returns>The numeric color value</returns>
    Private Function GetColorNumber(colorName As String) As Integer
        Select Case colorName.ToLower()
            Case "blue" : Return 0    ' PlayerBlue
            Case "red" : Return 1     ' PlayerRed
            Case "green" : Return 2   ' PlayerGreen
            Case "yellow" : Return 3  ' PlayerYellow
            Case "cyan" : Return 4    ' PlayerCyan
            Case "magenta" : Return 5 ' PlayerMagenta
            Case "black" : Return 6   ' PlayerBlack
            Case "white" : Return 7   ' Not defined in enum but used in game
            Case Else : Return 0      ' Default to blue
        End Select
    End Function

    ''' <summary>
    ''' Converts a direction string to its corresponding numeric value for Outpost 2
    ''' </summary>
    ''' <param name="direction">The direction string to convert</param>
    ''' <returns>The numeric direction value</returns>
    Private Function GetDirectionValue(direction As String) As Integer
        Select Case direction.ToLower()
            Case "east" : Return 0      ' East
            Case "southeast" : Return 1 ' SouthEast
            Case "south" : Return 2     ' South
            Case "southwest" : Return 3 ' SouthWest
            Case "west" : Return 4      ' West
            Case "northwest" : Return 5 ' NorthWest
            Case "north" : Return 6     ' North
            Case "northeast" : Return 7 ' NorthEast
            Case Else : Return 0        ' Default to East
        End Select
    End Function

    ''' <summary>
    ''' Converts a yield string to its corresponding numeric value for beacon creation
    ''' </summary>
    ''' <param name="yieldStr">The yield string to convert</param>
    ''' <returns>The numeric yield value</returns>
    Private Function GetYieldValue(yieldStr As String) As Integer
        Select Case yieldStr
            Case "Bar3" : Return 0
            Case "Bar2" : Return 1
            Case "Bar1" : Return 2
            Case "Random" : Return -1
            Case Else : Return -1 ' Default to random
        End Select
    End Function

    ''' <summary>
    ''' Converts a variant string to its corresponding numeric value for beacon creation
    ''' </summary>
    ''' <param name="variantStr">The variant string to convert</param>
    ''' <returns>The numeric variant value</returns>
    Private Function GetVariantValue(variantStr As String) As Integer
        Select Case variantStr
            Case "Variant1" : Return 0
            Case "Variant2" : Return 1
            Case "Variant3" : Return 2
            Case "Random" : Return -1
            Case Else : Return -1 ' Default to random
        End Select
    End Function

    ''' <summary>
    ''' Gets the human-readable name for a starship part based on its ID
    ''' </summary>
    ''' <param name="partId">The starship part ID</param>
    ''' <returns>The human-readable name of the starship part</returns>
    Private Function GetStarshipPartName(partId As Integer) As String
        Select Case partId
            Case 88 : Return "EDWARD Satellite"
            Case 89 : Return "Solar Satellite"
            Case 90 : Return "Ion Drive Module"
            Case 91 : Return "Fusion Drive Module"
            Case 92 : Return "Command Module"
            Case 93 : Return "Fueling Systems"
            Case 94 : Return "Habitat Ring"
            Case 95 : Return "Sensor Package"
            Case 96 : Return "Skydock"
            Case 97 : Return "Stasis Systems"
            Case 98 : Return "Orbital Package"
            Case 99 : Return "Phoenix Module"
            Case 100 : Return "Rare Metals Cargo"
            Case 101 : Return "Common Metals Cargo"
            Case 102 : Return "Food Cargo"
            Case 103 : Return "Evacuation Module"
            Case 104 : Return "Children Module"
            Case Else : Return $"Unknown Starship Part ({partId})"
        End Select
    End Function

#End Region

#Region "Helper Functions - Cargo Mapping"

    ''' <summary>
    ''' Maps cargo type from OPM format to the correct map_id enum value for Outpost 2
    ''' </summary>
    ''' <param name="cargoTypeStr">The cargo type string from OPM data</param>
    ''' <returns>The corresponding map_id enum value</returns>
    Private Function MapCargoType(cargoTypeStr As String) As String
        Select Case cargoTypeStr
        ' Resources
            Case "Food"
                Return "FoodCargo"
            Case "CommonMetal"
                Return "CommonMetalsCargo"
            Case "RareMetal"
                Return "RareMetalsCargo"
            Case "CommonOre"
                Return "CommonMetalsCargo" ' Best approximation
            Case "RareOre"
                Return "RareMetalsCargo"   ' Best approximation
            Case "CommonRubble"
                Return "CommonMetalsCargo" ' Best approximation 
            Case "RareRubble"
                Return "RareMetalsCargo"   ' Best approximation

        ' Weapons
            Case "Microwave"
                Return "Microwave"
            Case "Laser"
                Return "Laser"
            Case "EMP"
                Return "EMP"
            Case "RPG"
                Return "RPG"
            Case "Starflare"
                Return "Starflare"
            Case "Supernova"
                Return "Supernova"
            Case "ESG"
                Return "ESG"
            Case "AcidCloud"
                Return "AcidCloud"
            Case "RailGun"
                Return "RailGun"
            Case "Stickyfoam"
                Return "Stickyfoam"
            Case "ThorsHammer"
                Return "ThorsHammer"

        ' Special cases
            Case "Garbage"
                Return "None" ' No specific enum value for garbage
            Case "90"
                Return "None" ' Unknown value, default to None

                ' Default case - return the original string if it's a valid building/unit type
            Case Else
                If IsNumeric(cargoTypeStr) Then
                    Return "None" ' Handle numeric cargo types
                Else
                    Return cargoTypeStr ' Return original string for buildings, etc.
                End If
        End Select
    End Function

    ''' <summary>
    ''' Gets the appropriate truck cargo type for SetTruckCargo function calls
    ''' </summary>
    ''' <param name="cargoTypeStr">The cargo type string from OPM data</param>
    ''' <returns>The corresponding truck cargo type</returns>
    Private Function GetTruckCargoType(cargoTypeStr As String) As String
        Select Case cargoTypeStr
        ' Basic cargo types
            Case "Food"
                Return "truckFood"
            Case "CommonOre"
                Return "truckCommonOre"
            Case "RareOre"
                Return "truckRareOre"
            Case "CommonMetal"
                Return "truckCommonMetal"
            Case "RareMetal"
                Return "truckRareMetal"
            Case "CommonRubble"
                Return "truckCommonRubble"
            Case "RareRubble"
                Return "truckRareRubble"
            Case "Garbage"
                Return "truckGarbage"

        ' Special handling for Spaceport cargo
            Case "Spaceport"
                Return "truckSpaceport"

        ' Spaceship modules mapping (by name)
            Case "EDWARDSatellite"
                Return "truckSpaceport" ' ID 88
            Case "SolarSatellite"
                Return "truckSpaceport" ' ID 89
            Case "IonDriveModule"
                Return "truckSpaceport" ' ID 90
            Case "FusionDriveModule"
                Return "truckSpaceport" ' ID 91
            Case "CommandModule"
                Return "truckSpaceport" ' ID 92
            Case "FuelingSystems"
                Return "truckSpaceport" ' ID 93
            Case "HabitatRing"
                Return "truckSpaceport" ' ID 94
            Case "SensorPackage"
                Return "truckSpaceport" ' ID 95
            Case "Skydock"
                Return "truckSpaceport" ' ID 96
            Case "StasisSystems"
                Return "truckSpaceport" ' ID 97
            Case "OrbitalPackage"
                Return "truckSpaceport" ' ID 98
            Case "PhoenixModule"
                Return "truckSpaceport" ' ID 99

                ' Handle numeric cargo types (likely spaceship module IDs)
            Case Else
                If IsNumeric(cargoTypeStr) Then
                    Dim cargoId As Integer = Integer.Parse(cargoTypeStr)
                    If cargoId >= 88 AndAlso cargoId <= 104 Then
                        Return "truckSpaceport" ' Spaceship part IDs
                    Else
                        Return "truckEmpty" ' Default for unknown numeric values
                    End If
                Else
                    Return "truckEmpty" ' Default
                End If
        End Select
    End Function

#End Region

#Region "Helper Functions - Wall/Tube Processing"

    ''' <summary>
    ''' Groups individual wall/tube elements into continuous lines for more efficient C++ code generation
    ''' </summary>
    ''' <param name="wallTubes">List of individual wall/tube elements</param>
    ''' <returns>List of wall/tube lines that can be created with CreateTubeOrWallLine</returns>
    Private Function GroupWallTubesIntoLines(wallTubes As List(Of WallTube)) As List(Of WallTubeLine)
        Dim result As New List(Of WallTubeLine)
        Dim processedWallTubes As New HashSet(Of WallTube)

        ' Group horizontal lines
        For Each wallTube In wallTubes
            If processedWallTubes.Contains(wallTube) Then Continue For

            Dim line As New WallTubeLine()
            line.StartX = wallTube.Position.X
            line.StartY = wallTube.Position.Y
            line.EndX = wallTube.Position.X
            line.EndY = wallTube.Position.Y
            line.Type = wallTube.TypeID

            processedWallTubes.Add(wallTube)

            ' Look for adjacent horizontal wall tubes
            Dim extendRight As Boolean = True
            While extendRight
                Dim nextX As Integer = line.EndX + 1
                Dim nextWallTube = wallTubes.FirstOrDefault(Function(w) _
                Not processedWallTubes.Contains(w) AndAlso
                w.Position.Y = line.EndY AndAlso
                w.Position.X = nextX AndAlso
                w.TypeID = line.Type)

                If nextWallTube IsNot Nothing Then
                    line.EndX = nextWallTube.Position.X
                    processedWallTubes.Add(nextWallTube)
                Else
                    extendRight = False
                End If
            End While

            ' If we found a line with more than one element, add it
            If line.StartX <> line.EndX Then
                result.Add(line)
            End If
        Next

        ' Group vertical lines
        For Each wallTube In wallTubes
            If processedWallTubes.Contains(wallTube) Then Continue For

            Dim line As New WallTubeLine()
            line.StartX = wallTube.Position.X
            line.StartY = wallTube.Position.Y
            line.EndX = wallTube.Position.X
            line.EndY = wallTube.Position.Y
            line.Type = wallTube.TypeID

            processedWallTubes.Add(wallTube)

            ' Look for adjacent vertical wall tubes
            Dim extendDown As Boolean = True
            While extendDown
                Dim nextY As Integer = line.EndY + 1
                Dim nextWallTube = wallTubes.FirstOrDefault(Function(w) _
                Not processedWallTubes.Contains(w) AndAlso
                w.Position.X = line.EndX AndAlso
                w.Position.Y = nextY AndAlso
                w.TypeID = line.Type)

                If nextWallTube IsNot Nothing Then
                    line.EndY = nextWallTube.Position.Y
                    processedWallTubes.Add(nextWallTube)
                Else
                    extendDown = False
                End If
            End While

            ' If we found a line with more than one element, add it
            If line.StartY <> line.EndY Then
                result.Add(line)
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Gets the individual wall/tube elements that couldn't be grouped into lines
    ''' </summary>
    ''' <param name="allWallTubes">All wall/tube elements</param>
    ''' <param name="groupedLines">The lines that were already grouped</param>
    ''' <returns>List of individual wall/tube elements that need separate creation calls</returns>
    Private Function GetIndividualWallTubes(allWallTubes As List(Of WallTube), groupedLines As List(Of WallTubeLine)) As List(Of WallTube)
        Dim processedPositions As New HashSet(Of String)

        ' Add all positions from grouped lines
        For Each line In groupedLines
            Dim x As Integer = line.StartX
            While x <= line.EndX
                Dim y As Integer = line.StartY
                While y <= line.EndY
                    processedPositions.Add($"{x},{y}")
                    y += 1
                End While
                x += 1
            End While
        Next

        ' Return wall/tubes that aren't in processed positions
        Return allWallTubes.Where(Function(w) Not processedPositions.Contains($"{w.Position.X},{w.Position.Y}")).ToList()
    End Function

#End Region

#Region "Data Classes"

    ''' <summary>
    ''' Represents a line of wall or tube elements for efficient C++ code generation
    ''' </summary>
    Private Class WallTubeLine
        ''' <summary>
        ''' Gets or sets the starting X coordinate of the line
        ''' </summary>
        Public Property StartX As Integer

        ''' <summary>
        ''' Gets or sets the starting Y coordinate of the line
        ''' </summary>
        Public Property StartY As Integer

        ''' <summary>
        ''' Gets or sets the ending X coordinate of the line
        ''' </summary>
        Public Property EndX As Integer

        ''' <summary>
        ''' Gets or sets the ending Y coordinate of the line
        ''' </summary>
        Public Property EndY As Integer

        ''' <summary>
        ''' Gets or sets the type of wall or tube (e.g., "Tube", "Wall")
        ''' </summary>
        Public Property Type As String
    End Class

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click

    End Sub

#End Region

End Class