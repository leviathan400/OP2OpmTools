Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

' OP2OpmTools
' https://github.com/leviathan400/OP2OpmTools
'
' Open OP2MissionEditor .opm json file and
' export as C++ code, json or txt report
'
' Outpost 2: Divided Destiny is a real-time strategy video game released in 1997.

Public Class fMain

#Region "Fields and Properties"

    Public ApplicationName As String = "OP2OpmTools"
    Public Version As String = "0.2.0"

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
        Me.Icon = My.Resources.powerapps

        btnExportCpp.Enabled = False
        btnExportJson.Enabled = False
        btnExportTxt.Enabled = False
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
                    writer.WriteLine("// Generated by " & ApplicationName & " v" & Version)
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
                    writer.WriteLine("Generated by " & ApplicationName & " v" & Version)
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

#End Region

End Class