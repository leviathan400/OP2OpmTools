
' OPM file
'
' Units:
' vehicles, buildings
'
' Beacons:
' mines etc
'
Public Class LevelDetails
    Public Property LevelDescription As String
    Public Property MapName As String
    Public Property MaxTechLevel As Integer
    Public Property MissionType As String
    Public Property NumPlayers As Integer
    Public Property TechTreeName As String
    Public Property UnitOnlyMission As Boolean

    Public Overrides Function ToString() As String
        Return $"{MapName} - {MissionType} for {NumPlayers} players"
    End Function
End Class

Public Class Unit
    Public Property BarVariant As String
    Public Property BarYield As String
    Public Property CargoAmount As Integer
    Public Property CargoType As String
    Public Property CreateWall As Boolean
    Public Property Direction As String
    Public Property Health As Single
    Public Property ID As Integer
    Public Property IgnoreLayout As Boolean
    Public Property Lights As Boolean
    Public Property MaxTubes As Object
    Public Property MinDistance As Integer
    Public Property Position As Point
    Public Property SpawnDistance As Integer
    Public Property TypeID As String
    Public Property PlayerIndex As Integer
    Public Property PlayerColor As String

    Public Overrides Function ToString() As String
        Return $"{TypeID} at Position {Position}"
    End Function

End Class

Public Class Point

    Public Property X As Integer
    Public Property Y As Integer

    Public Overrides Function ToString() As String
        Return $"({X}, {Y})"
    End Function

End Class

Public Class WallTube

    Public Property Position As Point
    Public Property TypeID As String

    Public Overrides Function ToString() As String
        Return $"{TypeID} at Position {Position}"
    End Function

End Class

Public Class Beacon
    Public Property BarVariant As String
    Public Property BarYield As String
    Public Property ID As Integer
    Public Property MapID As String
    Public Property OreType As String
    Public Property Position As Point

    Public Overrides Function ToString() As String
        Return $"{MapID} ({OreType}) at {Position.X}, {Position.Y}"
    End Function
End Class