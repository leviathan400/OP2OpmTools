# The `.opm` File Format — Reference

The `.opm` (**O**utpost 2 **P**lugin **M**ission) file is the native save format of the
[OP2MissionEditor](https://github.com/leviathan400/OP2MissionEditor). It is a plain‑text
**JSON** document describing a complete Outpost 2 mission: the level metadata, every
player and their starting resources, all placed units / walls / tubes, map beacons,
markers, wreckage, music, and scripting structures (triggers, disasters and regions).

This document is a reference for the structure and field meanings of that JSON, compiled
from:

- Real sample files: `cnewmap.opm` and `cUnsettledEarth.opm` (editor / DataContract
  output), and SDK 3.0 `cSDKFloodPlain.opm` / `cSDKPieChart.opm` (Newtonsoft output).
- **The authoritative data model: the OP2DotNetMissionSDK source**, namespace
  `DotNetMissionReader`, under `DotNetMissionSDK/MissionReader/Json/`. Each `[DataMember]`
  attribute defines the exact JSON key; the C# property name may differ (e.g. the property
  `SdkVersion` serializes as `"SDKVersion"`). The class names below refer to these files.
  (The Unity editor bundles an older build of the same model under the
  `DotNetMissionSDK.Json` namespace.)
- The parser/exporter in [OP2OpmTools](https://github.com/leviathan400/OP2OpmTools)
  (`OP2OpmTools/Unit.vb`, `OP2OpmTools/fMain.vb`).

---

## 1. How it is produced and consumed

| | |
|---|---|
| **Writer** | OP2MissionEditor → `UserData.SaveMission()`. Uses .NET `System.Runtime.Serialization.Json.DataContractJsonSerializer` over the `MissionRoot` class, then pretty‑prints with `Utility.JsonFormatter` (tab indentation, newline after each value). |
| **Reader (editor)** | `DotNetMissionSDK.Json.MissionReader.GetMissionData(path)` → `DataContractJsonSerializer` → `MissionRoot`. |
| **Reader (OP2OpmTools)** | Newtonsoft.Json: `JObject.Parse()` for navigation, `JsonConvert.DeserializeObject(Of T)()` into the VB.NET model classes in `Unit.vb`. |

Two writer styles exist in the wild, and readers must accept both:

| Style | Indent | Member order | Default values | Produced by |
|---|---|---|---|---|
| **DataContract** | tabs | **alphabetical** | all members emitted (e.g. `"ID":0`) | OP2MissionEditor `DataContractJsonSerializer` (e.g. `cnewmap.opm`, `cUnsettledEarth.opm`) |
| **Newtonsoft** | 2 spaces | **declaration order** | defaults often omitted (`ID:0`, etc. dropped) | newer SDK 3.0 tooling (e.g. `cSDKFloodPlain.opm`, `cSDKPieChart.opm`) |

Key consequences:

- **Member names are PascalCase** (`LevelDetails`, `MapName`, `TypeID`, `CenterView`, …).
- **Member order is writer-dependent and not significant** — read by name, never by
  position. The DataContract writer sorts alphabetically; the Newtonsoft writer keeps
  declaration order.
- **Missing members default**, and writers may omit a member whose value is the type
  default. A reader must treat an absent field as its default (absent `ID` ⇒ `0`, etc.).
  Newer files routinely omit `ID:0` on beacons, markers and wreckage; the editor's
  DataContract output includes them. **Exception:** `UnitData` sets non‑type defaults on
  deserialize — an omitted `Health` ⇒ `1` and an omitted `Lights` ⇒ `true` (via the SDK's
  `OnDeserializing` hook). Constructor defaults for new missions (e.g. `MoraleLevel:"Good"`,
  `Kids:10`, `Workers:14`, `Scientists:8`, `Food:1000`) are **not** applied on read — those
  scalars deserialize to plain type defaults (`0`/`""`) when absent.
- **New/unknown members may appear** (e.g. `AIImpl` on a player — see [§5](#5-player)).
  Readers should ignore members they don't recognize rather than fail.
- `null` is written literally (e.g. `"Name":null`, `"BarVariant":null`).
- Enums are serialized **by name as strings** (e.g. `"Colony"`, `"Blue"`, `"Good"`),
  not by integer.
- Numeric fields may be written **without a decimal** when whole — `Health` appears as
  `1` (full) or `0.5`; treat it as a float regardless.

When the editor saves, it writes **three** sibling files: `mission.opm` (this format),
the `.map` file referenced by `MapName`, and a compiled `mission.dll` plugin.

---

## 2. Top‑level object (`MissionRoot`)

```jsonc
{
  "Disasters":      [ ... ],   // array of Disaster   — reserved (editor leaves empty)
  "LevelDetails":   { ... },   // object  — mission metadata
  "MasterVariant":  { ... },   // object  — base variant applied to all variants
  "MissionVariants":[ ... ],   // array of MissionVariant — alternate variants
  "Regions":        [ ... ],   // array of Region     — reserved (editor leaves empty)
  "SDKVersion":     "0",       // string  — format/SDK version tag
  "Triggers":       [ ... ]    // array of Trigger    — reserved (editor leaves empty)
}
```

| Field | Type | Notes |
|---|---|---|
| `SDKVersion` | string | Version tag of the SDK that wrote the file (e.g. `"0"`). Used by the editor and plugin exporter for compatibility. |
| `LevelDetails` | object | Mission‑wide metadata. See [§3](#3-leveldetails). |
| `MasterVariant` | object | The base [`Variant`](#4-variant-mastervariant--missionvariants) that always applies. Contains the players, the base `TethysGame`, difficulty overrides, and layouts. |
| `MissionVariants` | array | Zero or more additional `Variant` objects. At play time the engine picks one variant and **concatenates** it onto `MasterVariant`. Usually empty for simple missions. |
| `Disasters`, `Triggers`, `Regions` | arrays | Mission‑global scripting structures defined by the SDK but **not authored by OP2MissionEditor** — it always writes them as `[]`. See [§9](#9-reserved-scripting-structures). |

---

## 3. `LevelDetails`

Mission metadata. (Model: `LevelDetails`. Note the description key is `LevelDescription`,
mapping to the C# property `Description`.)

```jsonc
"LevelDetails": {
  "LevelDescription":"No description",
  "MapName":"newworld.map",
  "MaxTechLevel":12,
  "MissionType":"Colony",
  "NumPlayers":2,
  "TechTreeName":"MULTITEK.TXT",
  "UnitOnlyMission":false
}
```

| Field | Type | Meaning |
|---|---|---|
| `LevelDescription` | string | Human‑readable mission description shown in‑game. |
| `MapName` | string | Filename of the companion `.map` file (well/terrain data). |
| `MaxTechLevel` | int | Highest tech level reachable in this mission. |
| `MissionType` | string enum | Game mode. See [`MissionType`](#missiontype). |
| `NumPlayers` | int | Number of players the mission supports. |
| `TechTreeName` | string | Tech tree definition file (e.g. `MULTITEK.TXT`, `TECH.TXT`). |
| `UnitOnlyMission` | bool | If `true`, a "units only" scenario (no full colony economy). |

### `MissionType`

Serialized by name. Integer values are the Outpost 2 engine constants (negative).

| Name | Value | Editor filename prefix |
|---|---|---|
| `MultiLastOneStanding` | -1 | `ml<N>` |
| `MultiMidas` | -2 | `mm<N>` |
| `MultiResourceRace` | -3 | `mr<N>` |
| `MultiSpaceRace` | -4 | `mf<N>` |
| `MultiLandRush` | -5 | `mu<N>` |
| `Tutorial` | -6 | `t` |
| `AutoDemo` | -7 | `a` |
| `Colony` | -8 | `c` |

(`<N>` = `NumPlayers`. Source: `UserData.GetMissionTypePrefix()`.)

---

## 4. `Variant` (`MasterVariant` & `MissionVariants[]`)

Both `MasterVariant` and each entry of `MissionVariants` use the same shape.

```jsonc
"MasterVariant": {
  "AutoLayouts":         [ ... ],  // array of AutoLayout
  "Name":                null,     // string|null — variant display name
  "Players":             [ ... ],  // array of Player
  "TethysDifficulties":  [ ... ],  // array of GameData — per-difficulty overrides
  "TethysGame":          { ... }   // GameData — base game/map setup
}
```

| Field | Type | Meaning |
|---|---|---|
| `Name` | string \| null | Variant name. `null` / unused on the master variant; named like `"Variant 1"` on additional variants. |
| `Players` | array | One [`Player`](#5-player) per participant. |
| `TethysGame` | object | The base [`GameData`](#6-gamedata-tethysgame): lighting, music, beacons, markers, wreckage. |
| `TethysDifficulties` | array | Optional [`GameData`](#6-gamedata-tethysgame) overrides, one per difficulty level, concatenated onto `TethysGame` when that difficulty is selected. Empty when difficulties aren't used. |
| `AutoLayouts` | array | `AutoLayout` entries — automatic base‑layout templates (C# property `Layouts`). See [`AutoLayout`](#autolayout). Typically empty. |

> **Concatenation model:** the effective mission = `MasterVariant` + selected
> `MissionVariants[i]` + (optionally) the selected `TethysDifficulties[d]` /
> `Player.difficulties[d]`. Variants/difficulties contribute *additional* units,
> beacons, etc. on top of the master.

---

## 5. `Player`

One entry per player in `Variant.Players`. (Model: `PlayerData`.)

```jsonc
{
  "Allies":            [ ],        // array of int  — allied player IDs
  "BotType":           "None",     // string enum
  "AIImpl":            "AIv2",     // string enum — OPTIONAL, AI implementation
  "Color":             "Blue",     // string enum
  "DifficultyResources":[ ],       // array of Resources — per-difficulty overrides
  "ID":                0,          // int — player index
  "IsEden":            true,       // bool — Eden vs Plymouth colony
  "IsHuman":           true,       // bool — human vs AI
  "Resources":         { ... }     // Resources — base starting state
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Zero‑based player index (0–6). |
| `IsEden` | bool | `true` = Eden colony, `false` = Plymouth. |
| `IsHuman` | bool | `true` = human‑controlled, `false` = AI. |
| `Color` | string enum | Player color. See [`Color`](#color). |
| `BotType` | string enum | AI behavior preset. See [`BotType`](#bottype). |
| `AIImpl` | string | **Optional.** Selects which AI implementation to instantiate for this player slot. Empty or `"TechCor"` ⇒ the reference `BotPlayer`; `"AIv2"` ⇒ the alternative implementation. Default empty keeps every existing `.opm` working unchanged. Present only in newer SDK 3.0 files, typically on AI players. |
| `Allies` | array of int | IDs of players this player is allied with. |
| `Resources` | object | Base starting [`Resources`](#7-resources) (used when no difficulty is selected). |
| `DifficultyResources` | array | Per‑difficulty [`Resources`](#7-resources) overrides, parallel to `Variant.TethysDifficulties`. Empty when unused. |

### `Color`

Serialized by name. Integer is the engine color number (`SetColorNumber`).

| Name | Value |
|---|---|
| `Blue` | 0 |
| `Red` | 1 |
| `Green` | 2 |
| `Yellow` | 3 |
| `Cyan` | 4 |
| `Magenta` | 5 |
| `Black` | 6 |
| `White` | 7 |

(Source: `OP2OpmTools/fMain.vb` color mapping.)

### `BotType`

AI behavior preset. Confirmed values: `None` (no AI / human), `Balanced`,
`LaunchStarship` (AI plays toward a starship‑launch victory). The SDK enum also includes
the other standard Outpost 2 AI archetypes. `None` is used together with `IsHuman:true`.
When a bot type is set, the player may also carry an [`AIImpl`](#5-player) selecting the
AI implementation.

---

## 6. `GameData` (`TethysGame`)

Map/game‑level setup, attached to each variant. (Model:
`DotNetMissionSDK.Json.GameData`.)

```jsonc
"TethysGame": {
  "Beacons":             [ ... ],   // array of Beacon
  "DaylightEverywhere":  false,     // bool
  "DaylightMoves":       true,      // bool
  "InitialLightLevel":   0,         // int
  "Markers":             [ ... ],   // array of Marker
  "MusicPlayList":       { ... },   // MusicPlayList
  "WallTubes":           [ ... ],   // array of WallTube (optional, see note)
  "Triggers":            [ ... ],   // array of Trigger — reserved
  "Wreckage":            [ ... ]    // array of Wreckage
}
```

| Field | Type | Meaning |
|---|---|---|
| `DaylightEverywhere` | bool | If `true`, the whole map is permanently lit (no day/night terminator). |
| `DaylightMoves` | bool | If `true`, the daylight band moves across the map over time. |
| `InitialLightLevel` | int | Starting position/level of the daylight band. |
| `MusicPlayList` | object | Background music sequence. See [`MusicPlayList`](#musicplaylist). |
| `Beacons` | array | Mining beacons, fumaroles and magma vents. See [`Beacon`](#beacon). |
| `Markers` | array | Map markers / location flags. See [`Marker`](#marker). |
| `Wreckage` | array | Discoverable wreckage (tech to salvage). See [`Wreckage`](#wreckage). |
| `WallTubes` | array | Map‑level (player‑neutral) walls/tubes. See [`WallTube`](#walltube). **Version‑dependent:** present in some SDK 3.0 build output (e.g. `cSDKFloodPlain.opm`), but **not** a member of the canonical `GameData` in the current SDK source (where walls/tubes live only under each player's `Resources`). Treat as optional. |
| `Triggers` | array | Variant‑scoped triggers — SDK‑defined, left empty by the editor. See [§9](#9-reserved-scripting-structures). |

### `MusicPlayList`

```jsonc
"MusicPlayList": {
  "RepeatStartIndex":0,
  "SongIDs":[ 0 ]
}
```

| Field | Type | Meaning |
|---|---|---|
| `SongIDs` | array of int | Ordered list of song IDs (`SongID` enum values) to play. |
| `RepeatStartIndex` | int | Index into `SongIDs` to loop back to once the list finishes. |

### `Beacon`

A resource beacon, fumarole, or magma vent. (Model: `Unit.vb` `Beacon`.)

```jsonc
{
  "BarVariant":"Random",   // string enum — ore graphic variant
  "BarYield":"Random",     // string enum — ore yield amount
  "ID":0,                  // int
  "MapID":"MiningBeacon",  // string enum — beacon kind
  "OreType":"Common",      // string enum — ore type
  "Position":{ "X":0, "Y":0 }
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Unique beacon identifier. |
| `MapID` | string enum | Beacon kind: `MiningBeacon`, `Fumarole`, `MagmaVent`. |
| `OreType` | string enum | See [`OreType`](#oretype). Applies to mining beacons. |
| `BarYield` | string enum | See [`Yield`](#baryield). Per‑truckload yield rate. |
| `BarVariant` | string enum | See [`Variant`](#barvariant). Visual / yield‑decay variant. |
| `Position` | [`Point`](#point) | Map tile location. |

#### `OreType`
| Name | Value |
|---|---|
| `Random` | -1 |
| `Common` | 0 |
| `Rare` | 1 |

#### `BarYield`
| Name | Value |
|---|---|
| `Random` | -1 |
| `Bar3` | 0 |
| `Bar2` | 1 |
| `Bar1` | 2 |

#### `BarVariant`
| Name | Value |
|---|---|
| `Random` | -1 |
| `Variant1` | 0 |
| `Variant2` | 1 |
| `Variant3` | 2 |

(`BarYield` / `BarVariant` value mappings from `OP2OpmTools/fMain.vb`.)

### `Marker`

A map marker / location flag.

```jsonc
{
  "ID":0,                  // optional; omitted when 0
  "MarkerType":"Circle",   // string enum (SDK MarkerType)
  "Position":{ "X":1, "Y":1 }
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Unique marker identifier. Optional — omitted when `0`. |
| `MarkerType` | string enum | Marker kind (SDK `MarkerType`). Known value: `Circle`. Other marker graphics include `Beaker`, `DNA`, `Common`, `Rare` (from the editor's resource art set). |
| `Position` | [`Point`](#point) | Map tile location. |

### `Wreckage`

Discoverable wreckage carrying a tech / starship part.

```jsonc
{
  "ID":0,                  // optional; omitted when 0
  "TechID":8000,           // int — numeric map_id / tech id of the item
  "IsVisible":true,        // bool
  "Position":{ "X":2, "Y":2 }
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Unique wreckage identifier. Optional — omitted when `0`. |
| `TechID` | int | The unit/tech represented by the wreckage, as a **numeric** `map_id` / tech id (e.g. `8000`). Note this is an integer here, unlike `Unit.TypeID` which is a string enum. |
| `IsVisible` | bool | Whether the wreckage is visible before discovery. |
| `Position` | [`Point`](#point) | Map tile location. |

---

## 7. `Resources`

Per‑player starting economy plus the list of units, walls and tubes the player begins
with. Appears as `Player.Resources` and inside `Player.DifficultyResources[]`.
(Model: `DotNetMissionSDK.Json.PlayerData.ResourceData`.)

```jsonc
"Resources": {
  "CenterView":        { "X":0, "Y":0 },  // Point — initial camera focus
  "CommonOre":         0,
  "CompletedResearch": [ ],               // array of int — researched tech IDs
  "Food":              1000,
  "FreeMorale":        false,             // bool — morale locked free of penalties
  "Kids":              10,
  "MoraleLevel":       "Good",            // string enum
  "RareOre":           0,
  "Scientists":        8,
  "SolarSatellites":   0,
  "TechLevel":         0,
  "Triggers":          [ ],               // array of Trigger — reserved
  "Units":             [ ... ],           // array of Unit
  "WallTubes":         [ ... ],           // array of WallTube
  "Workers":           14
}
```

| Field | Type | Meaning |
|---|---|---|
| `CenterView` | [`Point`](#point) | Tile the player's camera centers on at start. |
| `CommonOre` | int | Starting common ore stockpile. |
| `RareOre` | int | Starting rare ore stockpile. |
| `Food` | int | Starting food stockpile. |
| `Kids` | int | Starting children population. |
| `Workers` | int | Starting worker population. |
| `Scientists` | int | Starting scientist population. |
| `SolarSatellites` | int | Number of solar power satellites in orbit. |
| `TechLevel` | int | Starting technology level. |
| `CompletedResearch` | array of int | Tech IDs already researched at start. |
| `MoraleLevel` | string enum | See [`MoraleLevel`](#moralelevel). |
| `FreeMorale` | bool | If `true`, morale is locked at the set level (no demand penalties). |
| `Units` | array | Pre‑placed units & structures. See [`Unit`](#8-unit). |
| `WallTubes` | array | Pre‑placed walls & tubes. See [`WallTube`](#walltube). |
| `Triggers` | array | Player‑scoped triggers — reserved, left empty by the editor. |

### `MoraleLevel`

Serialized by name. Standard Outpost 2 morale tiers: `Excellent`, `Good`, `Fair`,
`Poor`, `Terrible` (sample uses `"Good"`).

---

## 8. `Unit`

A pre‑placed vehicle or structure. (Model: `OP2OpmTools/Unit.vb` `Unit`.) Not every
field is meaningful for every unit type; the editor writes the ones relevant to the
unit it placed.

```jsonc
{
  "ID":0,
  "TypeID":"CargoTruck",      // string enum (map_id) — unit/structure type
  "Position":{ "X":13, "Y":8 },
  "Direction":"East",         // string enum
  "Health":1.0,               // float 0..1
  "Lights":true,              // bool — vehicle headlights
  "CargoType":"None",         // string enum — cargo / weapon / module
  "CargoAmount":0,            // int — quantity or module/cargo id
  "BarVariant":"Random",      // string enum — for ore mines
  "BarYield":"Random",        // string enum — for ore mines
  "CreateWall":false,         // bool
  "IgnoreLayout":false,       // bool
  "MaxTubes":null,            // int|null
  "MinDistance":0,            // int
  "SpawnDistance":0           // int
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Unique unit identifier. |
| `TypeID` | string enum | Unit/structure type (`map_id`, e.g. `CargoTruck`, `Tiger`, `CommandCenter`). See [§10](#10-mapid-unitstructure-type-names). |
| `Position` | [`Point`](#point) | Map tile location. |
| `Direction` | string enum | Facing direction (vehicles). See [`Direction`](#direction). |
| `Health` | float | Health as a fraction `0.0`–`1.0` (1.0 = full). |
| `Lights` | bool | Headlights on/off (vehicles). |
| `CargoType` | string enum | Cargo for trucks, **weapon** for combat units, or starship module type. See [cargo/weapon tables](#cargo--weapons). |
| `CargoAmount` | int | Cargo quantity, weapon‑related amount, or starship module ID (see [Starship modules](#starship-module-ids)). |
| `BarVariant` | string enum | Ore variant for ore‑mine structures (see [`BarVariant`](#barvariant)). |
| `BarYield` | string enum | Ore yield for ore‑mine structures (see [`BarYield`](#baryield)). |
| `CreateWall` | bool | Whether a wall is created with/around the unit. |
| `IgnoreLayout` | bool | Skip auto‑layout placement rules for this unit. |
| `MaxTubes` | int \| null | Max tube connections (structures); `null` if unspecified. |
| `MinDistance` | int | Minimum placement distance constraint. |
| `SpawnDistance` | int | Spawn distance constraint. |

### `Direction`

Serialized by name (the editor's canonical casing is camelCase compound names, e.g.
`SouthEast`). Reads are **case-insensitive** — `Southeast` also parses — but emit the
canonical form for a clean round-trip. Integer is the engine facing index.

| Name | Value |
|---|---|
| `East` | 0 |
| `SouthEast` | 1 |
| `South` | 2 |
| `SouthWest` | 3 |
| `West` | 4 |
| `NorthWest` | 5 |
| `North` | 6 |
| `NorthEast` | 7 |

### Cargo & weapons

**Truck cargo** (`CargoType` on `CargoTruck`). The editor serialises this from the SDK's `TruckCargo`
enum **by name** (`((TruckCargo)value).ToString()`), so the string is the enum member name — *not*
always the obvious cargo name. Two are easy to misread, and one cargo isn't named at all:

| `CargoType` (in `.opm`) | TruckCargo value | Means | OP2 SDK constant | OP2Lua `cargo` name |
|---|---|---|---|---|
| `Food` | 1 | food | `truckFood` | `Food` |
| `CommonOre` | 2 | common ore | `truckCommonOre` | `CommonOre` |
| `RareOre` | 3 | rare ore | `truckRareOre` | `RareOre` |
| `CommonMetal` | 4 | common metal | `truckCommonMetal` | `CommonMetal` |
| `RareMetal` | 5 | rare metal | `truckRareMetal` | `RareMetal` |
| `CommonRubble` | 6 | common rubble | `truckCommonRubble` | `CommonRubble` |
| `RareRubble` | 7 | rare rubble | `truckRareRubble` | `RareRubble` |
| `Spaceport` | 8 | **starship module** (the module `map_id` is in `CargoAmount`, 88–104) | `truckSpaceport` | `Spacecraft` |
| `Garbage` | 9 | **Wreckage** (the SDK enum literally comments this value as "Wreckage") | `truckGarbage` | `Wreckage` |
| `90` *(or `10`)* | 10 | **GeneBank** — the SDK `TruckCargo` enum has **no name** for value 10, so the editor writes the raw `map_id` (`mapGeneBank` = 90) | *(none; use `(TruckCargo)10`)* | `GeneBank` |

> **Gotchas for converters:**
> - `Garbage` is **Wreckage**, not literal garbage. `Spaceport` is a **starship module** (read the
>   module id from `CargoAmount`, not `CargoType`).
> - A **numeric** `CargoType` (the editor only ever writes `"90"` here) is a **GeneBank** — it is *not*
>   a starship module. Modules always come through as `CargoType = "Spaceport"` with the id in
>   `CargoAmount`. (OP2OpmTools handles all of these — see `GetTruckCargoType` / `LuaTruckCargo`.)
> - OP2Lua uses the **Tethys** `CargoType` names (`Spacecraft`/`Wreckage`/`GeneBank`), which differ
>   from the SDK names above, so the Lua export translates them.

**Weapons** (`CargoType` on combat units — `Lynx`, `Panther`, `Tiger`, `Spider`,
`Scorpion`, `GuardPost`): `Microwave`, `Laser`, `EMP`, `RPG`, `Starflare`, `Supernova`,
`ESG`, `AcidCloud`, `RailGun`, `Stickyfoam`, `ThorsHammer` (mapped to the corresponding
`map_id` weapon, e.g. `mapMicrowave`).

##### Weapon × chassis validity (the editor only instantiates art-backed combos)

OP2MissionEditor renders OP2's actual unit art, so a `CargoType` weapon is only valid on a
chassis that has a turret graphic for it. Placing an unarted combo loads but **throws
`ArgumentException: The Object you want to instantiate is null` during "Creating units."**
The combat **bomb** weapons — **`Starflare` and `Supernova`** — only have turret art on the
**Panther and Tiger** chassis; they are **not** valid on `Lynx`, `Spider`, `Scorpion`, or
`GuardPost`. The other nine weapons (`Laser`, `Microwave`, `RailGun`, `EMP`, `RPG`, `ESG`,
`AcidCloud`, `Stickyfoam`, `ThorsHammer`) are accepted on all six combat-unit types.

| Weapon | Lynx | Panther | Tiger | Spider | Scorpion | GuardPost |
|---|:--:|:--:|:--:|:--:|:--:|:--:|
| Laser / Microwave / RailGun / EMP / RPG / ESG / AcidCloud / Stickyfoam / ThorsHammer | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| **Starflare**, **Supernova** (bombs) | ❌ | ✅ | ✅ | ❌ | ❌ | ❌ |

(Empirical, from loading generated missions in OP2MissionEditor — `Tiger`+bomb confirmed by
an editor round-trip save. Reflects which combos OP2 has *art* for, not necessarily the
in-game tech-tree build rules.)

#### Starship module IDs

When a `CargoTruck`/ConVec carries a starship part, `CargoAmount` holds the module ID:

| ID | Module | ID | Module |
|---|---|---|---|
| 88 | EDWARD Satellite | 97 | Stasis Systems |
| 89 | Solar Satellite | 98 | Orbital Package |
| 90 | Ion Drive Module | 99 | Phoenix Module |
| 91 | Fusion Drive Module | 100 | Rare Metals Cargo |
| 92 | Command Module | 101 | Common Metals Cargo |
| 93 | Fueling Systems | 102 | Food Cargo |
| 94 | Habitat Ring | 103 | Evacuation Module |
| 95 | Sensor Package | 104 | Children Module |
| 96 | Skydock | | |

---

## WallTube

A single wall or tube tile. Appears in `Resources.WallTubes` (per‑player) and, in some
SDK 3.0 builds, in `GameData.WallTubes` (map‑level). (Model: `WallTubeData`.)

```jsonc
{
  "TypeID":"Tube",          // string enum (map_id)
  "Position":{ "X":19, "Y":14 }   // canonical key; some builds emit "Location"
}
```

| Field | Type | Meaning |
|---|---|---|
| `TypeID` | string enum | One of `Tube`, `Wall`, `LavaWall`, `MicrobeWall`. |
| `Position` | [`Point`](#point) | Map tile location. |

> **Coordinate key drift:** the canonical `WallTubeData` model names this member
> **`Position`** (like every other placed object). However, the shipped SDK 3.0 demo files
> (`cSDKFloodPlain.opm`, `cSDKPieChart.opm`) emit it as **`Location`** in their
> `GameData.WallTubes` arrays. A robust reader should accept **either** key. Note
> `OP2OpmTools/Unit.vb` models it as `Position`, so the `Location` variant may not
> round‑trip through that tool.
>
> OP2OpmTools groups runs of adjacent same‑type wall/tube tiles into
> `CreateTubeOrWallLine(...)` calls when exporting C++.

## Point

A 2‑D tile coordinate. SDK type: `DataLocation` (a struct). Used by every `Position`,
`CenterView`, `AutoLayout.BaseCenterPt`, and the `Location` variant of `WallTube`.

```jsonc
{ "X":13, "Y":8 }
```

| Field | Type |
|---|---|
| `X` | int |
| `Y` | int |

> **Coordinate note:** `.opm` stores logical map coordinates. When OP2OpmTools generates
> engine C++ it applies the offset `LOCATION(X + 31, Y - 1)` to convert editor
> coordinates to in‑game world coordinates (the map's playable region is inset by the
> 31‑tile border). The raw `.opm` values are *pre‑offset* editor coordinates.

## DataRect

An axis‑aligned tile rectangle. SDK type: `DataRect`. Used by `Disaster` (`SrcRect`,
`DestRect`) and `Region` (`Rect`).

```jsonc
{ "MinX":0, "MinY":0, "MaxX":0, "MaxY":0 }
```

| Field | Type |
|---|---|
| `MinX`, `MinY` | int (top‑left tile) |
| `MaxX`, `MaxY` | int (bottom‑right tile) |

## AutoLayout

A template that auto‑places a base layout for a player. Lives in `Variant.AutoLayouts`.
(Model: `AutoLayout`; C# property `Layouts`.) Usually empty.

```jsonc
{
  "PlayerID":0,
  "BaseCenterPt":{ "X":0, "Y":0 },
  "Units":[ /* array of Unit, see §8 */ ]
}
```

| Field | Type | Meaning |
|---|---|---|
| `PlayerID` | int | Player the layout belongs to. |
| `BaseCenterPt` | [`Point`](#point) | Anchor tile the layout is centered on. |
| `Units` | array | [`Unit`](#8-unit) entries placed relative to the base. The unit fields `IgnoreLayout`, `MinDistance`, `SpawnDistance`, `CreateWall`, `MaxTubes` feed the layout algorithm. |

---

## 9. Scripting structures: Disasters, Regions, Triggers

These drive mission logic. OP2MissionEditor has no UI to author them, so editor‑written
files leave them as empty arrays `[]` — but they are **fully defined in the SDK model**
and may be hand‑authored or produced by other tooling. All enum‑typed members are
serialized **by name as strings**.

> **Two distinct trigger schemas exist.** They are *not* interchangeable:
> - `MissionRoot.Triggers` → array of **`OP2TriggerData`** (flat, classic "OP2 trigger").
> - `GameData.Triggers` and `Resources.Triggers` → array of **`TriggerData`** (the newer
>   condition/action model).

### 9.1 `Disaster` (`MissionRoot.Disasters[]`)

Model: `DisasterData`.

```jsonc
{
  "ID":0,
  "Type":"Meteor",                       // DisasterType enum
  "SrcRect":{ "MinX":0,"MinY":0,"MaxX":0,"MaxY":0 },   // DataRect — spawn area
  "DestRect":{ "MinX":0,"MinY":0,"MaxX":0,"MaxY":0 },  // DataRect — target area (moving disasters)
  "Size":0,                              // magnitude / size / spread speed
  "Duration":0,
  "StartTime":0,
  "EndTime":0,
  "MinDelay":0,
  "MaxDelay":0
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Identifier. |
| `Type` | string enum | See [`DisasterType`](#disastertype). |
| `SrcRect` | [`DataRect`](#datarect) | Source/spawn rectangle. |
| `DestRect` | [`DataRect`](#datarect) | Destination rectangle (for travelling disasters like vortex/storm). |
| `Size` | int | Magnitude (meteor size, quake magnitude, spread speed, etc.). |
| `Duration` | int | Disaster duration (ticks). |
| `StartTime`, `EndTime` | int | Active time window (ticks). |
| `MinDelay`, `MaxDelay` | int | Random delay range between recurrences. |

#### `DisasterType`
`Meteor`, `Earthquake`, `Lightning`, `Tornado`, `Eruption`, `Blight`.

### 9.2 `Region` (`MissionRoot.Regions[]`)

Model: `RegionData`. A named rectangle referenced by triggers.

```jsonc
{
  "Name":"",
  "Rect":{ "MinX":0,"MinY":0,"MaxX":0,"MaxY":0 }
}
```

| Field | Type | Meaning |
|---|---|---|
| `Name` | string | Region name (referenced by trigger condition/action region fields). |
| `Rect` | [`DataRect`](#datarect) | The region's tile bounds. |

### 9.3 `OP2TriggerData` (`MissionRoot.Triggers[]`)

The classic flat trigger. Most members are only meaningful for specific `Type` values.

```jsonc
{
  "ID":0,
  "Type":"Victory",            // TriggerType enum
  "Enabled":true,
  "OneShot":false,
  "TriggerID":0,
  "Message":"",
  "PlayerID":0,
  "Count":0,
  "CompareType":"",            // TriggerCompare enum (as string)
  "UnitType":"",               // map_id (as string)
  "Time":0, "MinTime":0, "MaxTime":0,
  "TechID":0,
  "X":0, "Y":0, "Width":0, "Height":0,
  "CargoOrWeaponType":"",
  "ResourceType":"",
  "CargoType":"",
  "CargoAmount":0,
  "ChildTriggerIDs":[ ],       // for "Set" triggers: child trigger IDs
  "NeededTriggers":0,          // min children that must fire (0 ⇒ all)
  "Actions":[ /* ActionData */ ]
}
```

| Field | Type | Notes |
|---|---|---|
| `ID` | int | Trigger identifier. |
| `Type` | string enum | See [`TriggerType`](#triggertype). |
| `Enabled` | bool | Whether the trigger starts active. |
| `OneShot` | bool | If `true`, fires once then disables. |
| `TriggerID` | int | Engine trigger handle / link id. |
| `Message` | string | Associated message text. |
| `PlayerID` | int | Subject player. |
| `Count` | int | Count threshold (for count/vehicle/building triggers). |
| `CompareType` | string enum | Comparison, see [`TriggerCompare`](#triggercompare). |
| `UnitType` | string | `map_id` name. |
| `Time`, `MinTime`, `MaxTime` | int | Time / time‑range parameters (ticks). |
| `TechID` | int | Research/tech id (research triggers). |
| `X`, `Y`, `Width`, `Height` | int | Rect/point parameters for `Point`/`Rect` triggers. |
| `CargoOrWeaponType`, `ResourceType`, `CargoType` | string | Cargo/weapon/resource selectors. |
| `CargoAmount` | int | Cargo quantity. |
| `ChildTriggerIDs` | int[] | For `Set` triggers: IDs of child triggers (1..7). |
| `NeededTriggers` | int | Minimum number of children that must fire; defaults to all. |
| `Actions` | array | [`ActionData`](#actiondata) to run when the trigger fires. |

#### `TriggerType`
`None`, `Victory`, `Failure`, `OnePlayerLeft`, `Evac`, `Midas`, `Operational`,
`Research`, `Resource`, `Kit`, `Escape`, `Count`, `VehicleCount`, `BuildingCount`,
`Attacked`, `Damaged`, `Time`, `TimeRange`, `Point`, `Rect`, `SpecialTarget`, `Set`.

#### `ActionData`
Model: `ActionData`.

```jsonc
{ "Type":"", "Message":"", "PlayerID":0, "SoundID":0 }
```

| Field | Type | Meaning |
|---|---|---|
| `Type` | string | Action type. |
| `Message` | string | Message text to display. |
| `PlayerID` | int | Target player. |
| `SoundID` | int | Sound to play. |

### 9.4 `TriggerData` (`GameData.Triggers[]`, `Resources.Triggers[]`)

The condition/action trigger model.

```jsonc
{
  "ID":0,
  "Enabled":true,
  "EventType":"OnGameTick",    // TriggerEventType enum
  "Condition":[ /* TriggerConditionData — NOTE: key is singular "Condition" */ ],
  "Actions":[ /* TriggerActionData */ ]
}
```

| Field | Type | Meaning |
|---|---|---|
| `ID` | int | Trigger identifier. |
| `Enabled` | bool | Whether the trigger starts active (defaults `true`). |
| `EventType` | string enum | When the trigger is evaluated. See [`TriggerEventType`](#triggereventtype). |
| `Condition` | array | List of [`TriggerConditionData`](#triggerconditiondata). **The JSON key is the singular `Condition`** even though it holds a list. |
| `Actions` | array | List of [`TriggerActionData`](#triggeractiondata). |

#### `TriggerEventType`
`CheckVictoryCondition`, `CheckDefeatCondition`, `OnGameMark`, `OnGameTick`,
`OnPlayerDefeated`, `OnVehicleEnteredRegion`, `OnVehicleInRegion`,
`OnVehicleExitedRegion`, `OnVehicleCreated`, `OnVehicleDestroyed`,
`OnStructureEnteredRegion`, `OnStructureInRegion`, `OnStructureExitedRegion`,
`OnStructureCreated`, `OnStructureDestroyed`, `OnStructureOperational`,
`OnStructureInoperable`, `OnStructureKitEnteredBay`, `OnResearchStarted` (=22),
`OnResearchCompleted`, `OnStarshipModuleDeployed`, `OnWreckageDiscovered`,
`OnWreckageLoaded`, `OnConvecLoaded`.

#### `TriggerConditionData`
Model: `TriggerConditionData`. The meaning of `Value`/`Value2`/`Value3` depends on `Type`.

```jsonc
{
  "Type":"PlayerKids",     // TriggerConditionType enum
  "Subject":"",            // player/unit category, Switch#, etc.
  "Comparison":"",         // TriggerCompare enum
  "Value":"", "Value2":"", "Value3":""
}
```

| Field | Type | Meaning |
|---|---|---|
| `Type` | string enum | See [`TriggerConditionType`](#triggerconditiontype). |
| `Subject` | string | Subject selector — a [`TriggerPlayerCategory`](#trigger-category-enums) / [`TriggerUnitCategory`](#trigger-category-enums) (name or non‑negative ID), or `Switch#`. |
| `Comparison` | string enum | See [`TriggerCompare`](#triggercompare). |
| `Value`, `Value2`, `Value3` | string | Operands; interpreted per `Type` (quantity, TopicID, colony, difficulty, morale, region, `map_id`, etc.). |

#### `TriggerActionData`
Model: `TriggerActionData`. The meaning of the `Value*` operands depends on `Type`.

```jsonc
{
  "Type":"SetPlayerKids",  // TriggerActionType enum
  "Modifier":"Set",        // TriggerModifier enum (Set/Add/Subtract)
  "Subject":"",            // category / map_id / Switch#, etc.
  "Subject2":"",           // e.g. map_id weapon
  "SubjectPlayer":"",      // TriggerPlayerCategory
  "SubjectRegion":"",      // Region name/id
  "SubjectQuantity":0,
  "Value":"", "Value2":"", "Value3":"", "Value4":"", "Value5":""
}
```

| Field | Type | Meaning |
|---|---|---|
| `Type` | string enum | See [`TriggerActionType`](#triggeractiontype). |
| `Modifier` | string enum | `Set`, `Add`, or `Subtract`. |
| `Subject` | string | Primary subject (category / `map_id` / `Switch#` depending on `Type`). |
| `Subject2` | string | Secondary subject (e.g. weapon `map_id`). |
| `SubjectPlayer` | string | Player category/id the action targets. |
| `SubjectRegion` | string | Region name/id the action operates in. |
| `SubjectQuantity` | int | Count of units/items the action affects. |
| `Value`…`Value5` | string | Operands; interpreted per `Type`. |

#### `TriggerCompare`
`Equals`, `NotEquals`, `GreaterThan`, `GreaterThanOrEqual`, `LessThan`,
`LessThanOrEqual`.

#### Trigger category enums
Used by `Subject` / `SubjectPlayer` / region fields. **Non‑negative values are taken as a
literal ID** (PlayerID / UnitID / RegionID); the named members are the negative
"special" selectors:

- **`TriggerPlayerCategory`** — `Any`(-1), `All`(-2), `TriggerOwner`(-3), `EventPlayer`(-4),
  `OwnerAllies`(-5), `OwnerEnemies`(-6), `EventAllies`(-7), `EventEnemies`(-8).
- **`TriggerUnitCategory`** — `EventUnit`(-1).
- **`TriggerRegion`** — `Anywhere`(-1).
- **`TriggerColonyType`** — `Eden`, `Plymouth`.

#### `TriggerConditionType`
Grouped by numeric range (game state `0+`, player `1000+`, unit `2000+`, beacon `3000+`):

> `CurrentMark`, `CurrentTick`, `CurrentRegion`, `CurrentTopic`, `SwitchState`,
> `PlayerEquals`, `PlayerDifficulty`, `PlayerColony`, `PlayerKids`, `PlayerWorkers`,
> `PlayerScientists`, `PlayerPopulation`, `PlayerCommonMetal`, `PlayerRareMetal`,
> `PlayerMaxCommonMetal`, `PlayerMaxRareMetal`, `PlayerFood`, `PlayerVehicleCount`,
> `PlayerUnitsInRegion`, `PlayerConvecsWithKitInRegion`, `PlayerTrucksWithCargoInRegion`,
> `PlayerMorale`, `PlayerResearch`, `PlayerAvailableWorkers`, `PlayerAvailableScientists`,
> `PlayerPowerGenerated`, `PlayerInactivePowerCapacity`, `PlayerPowerConsumed`,
> `PlayerPowerAvailable`, `PlayerIdleStructureCount`, `PlayerActiveStructureCount`,
> `PlayerStructureCount`, `PlayerUnpoweredStructureCount`, `PlayerWorkersRequired`,
> `PlayerScientistsRequired`, `PlayerScientistsAsWorkers`, `PlayerScientistsResearching`,
> `PlayerFoodProduced`, `PlayerFoodConsumed`, `PlayerFoodLacking`,
> `PlayerNetFoodProduction`, `PlayerSolarSatelliteCount`, `PlayerRecreationCapacity`,
> `PlayerForumCapacity`, `PlayerMedicalCenterCapacity`, `PlayerResidenceCapacity`,
> `PlayerCanAccumulateOre`, `PlayerCanAccumulateRareOre`, `PlayerCanBuildSpaceport`,
> `PlayerCanBuildStructure`, `PlayerCanBuildVehicle`, `PlayerCanResearchTopic`,
> `PlayerHasVehicle`, `PlayerHasActiveCommandCenter`, `PlayerStarshipModuleDeployed`,
> `PlayerStarshipProgress`, `UnitID`, `UnitOwnerID`, `UnitType`, `UnitBusy`, `UnitEMPed`,
> `StructureEnabled`, `StructureDisabled`, `UnitPositionX`, `UnitPositionY`, `UnitHealth`,
> `ConvecKitType`, `TruckCargoType`, `TruckCargoQuantity`, `UnitWeapon`, `UnitLastCommand`,
> `UnitCurrentAction`, `VehicleStickyfoamed`, `VehicleESGed`,
> `UniversityWorkersInTraining`, `SpaceportLaunchpadCargo`, `UnitLights`, `StructurePower`,
> `StructureWorkers`, `StructureScientists`, `StructureInfected`, `LabResearchTopic`,
> `LabScientistCount`, `UnitHasWeapon`, `StructureHasEmptyBay`, `StructureHasBayWithCargo`,
> `UnitInRegion`, `BeaconTruckLoadsSoFar`, `BeaconSurveyedByPlayer`.

#### `TriggerActionType`
Grouped by numeric range (global `0+`, player/map `1000+`, units‑in‑region `2000+`,
single‑unit `3000+`, create `4000+`):

> `WaitTicks`, `PreserveTrigger`, `MoveRegionToUnit`, `MoveRegionToPosition`,
> `MoveRegionToRegion`, `MoveRegionToUnitTypeInRegion`, `SetRegionSize`, `SetSwitch`,
> `SetPlayerName`, `SetPlayerColor`, `SetPlayerBotType`, `SetPlayerDefeated`,
> `SetPlayerRLVs`, `SetPlayerSolarSatellites`, `SetPlayerEDWARDSatellites`,
> `SetPlayerKids`, `SetPlayerWorkers`, `SetPlayerScientists`, `SetPlayerCommonMetal`,
> `SetPlayerRareMetal`, `SetPlayerFood`, `SetPlayerTechLevel`, `StarvePlayerPopulation`,
> `SetPlayerResearchCompleted`, `SetPlayerAlliance`, `CapturePlayerRLV`,
> `SetPlayerCameraToRegion`, `ShowMessageToPlayerAtRegion`, `ShowMessageToPlayerForUnit`,
> `ShowMessageToPlayerFromPlayer`, `SetDaylightEverywhere`, `SetDaylightMoves`,
> `SetPlayerMorale`, `FreePlayerMorale`, `SetRandomSeed`, `CreateMeteor`,
> `CreateEarthquake`, `CreateStorm`, `CreateVortex`, `CreateVolcano`, `CreateMicrobe`,
> `SetMicrobeSpreadSpeed`, `RemoveMicrobe`, `FireMissileEMP`, `SetMusic`, `SetTileIndex`,
> `SetCellType`, `SetLavaPossible`, `SetLightLevel`, `SetUnitsInRegionToAttackEnemyUnits`,
> `SetUnitsInRegionToAttackGround`, `DeployMinesInRegion`, `SetUnitsInRegionToBulldoze`,
> `SetUnitsInRegionToDock`, `SetUnitsInRegionToStandGround`,
> `SetEarthworkersInRegionToBuildWallOrTube`, `SetEarthworkersInRegionToRemoveWallOrTube`,
> `SetFactoriesInRegionToProduceUnit`, `SetStructuresInRegionToTransferBayCargo`,
> `SetStructuresInRegionToLoadCargo`, `SetStructuresInRegionToUnloadCargo`,
> `SetTrucksInRegionToDumpCargo`, `SetLabsInRegionToResearch`,
> `SetUniversitiesInRegionToTrain`, `SetUnitsInRegionToRepair`,
> `SetSpidersInRegionToReprogram`, `SetConvecsInRegionToDismantle`,
> `SetTrucksInRegionToSalvage`, `SetUnitsInRegionToGuard`, `SetUnitsInRegionToPoof`,
> `SetSpaceportsInRegionToTransferLaunchpadCargo`, `SetHealthForUnitsInRegion`,
> `SetAutotargetForUnitsInRegion`, `KillUnitsInRegion`, `SetUnitsInRegionToSelfDestruct`,
> `TransferUnitsInRegionToPlayer`, `SetWeaponForUnitsInRegion`, `SetLightsForUnitsInRegion`,
> `SetUnitsInRegionToMove`, `TeleportUnitsInRegion`, `SetConvecsInRegionToBuild`,
> `SetCargoForConvecsInRegion`, `SetCargoForTrucksInRegion`, `SetStructuresInRegionToIdle`,
> `SetStructuresInRegionToActivate`, `SetStructuresInRegionToStop`,
> `SetStructuresInRegionToInfected`, `SetSpaceportsInRegionToLaunch`,
> `SetStructuresInRegionBayCargo`, `SetFactoriesInRegionToDevelop`,
> `SetUnitsInRegionToClearSpecialTarget`, and the single‑unit (`3000+`) equivalents
> (`SetUnitToAttackEnemyUnits`, `DeployMine`, `SetEarthworkerToBuildWallOrTube`,
> `SetFactoryToProduceUnit`, `KillUnit`, `SetCargoForTruck`, `SetFactoryToDevelop`, …),
> and the create actions (`4000+`): `CreateCargoTruck`, `CreateConvec`, `CreateVehicle`,
> `CreateStructure`, `CreateMine`, `CreateBeacon`, `CreateMarker`, `CreateWreckage`,
> `CreateWallOrTube`.

> If you author triggers by hand, follow the conventions: PascalCase member names, enums
> by name, category fields as either a special name or a literal non‑negative ID.

---

## 10. `map_id` (unit/structure type names)

`TypeID` (units, walls/tubes) and beacon `MapID` use the Outpost 2 `map_id`
enum, serialized as the **bare name string** (the engine constant prepends `map`, e.g.
`CommandCenter` → `mapCommandCenter`). Note `Wreckage.TechID` is the exception — it stores
the **numeric** `map_id` value (e.g. `8000`) rather than the name.

**Structures** (validated by OP2OpmTools):
`CommonOreMine`, `RareOreMine`, `GuardPost`, `LightTower`, `CommonStorage`,
`RareStorage`, `Forum`, `CommandCenter`, `MHDGenerator`, `Residence`, `RobotCommand`,
`TradeCenter`, `BasicLab`, `MedicalCenter`, `Nursery`, `SolarPowerArray`,
`RecreationFacility`, `University`, `Agridome`, `DIRT`, `Garage`, `MagmaWell`,
`MeteorDefense`, `GeothermalPlant`, `ArachnidFactory`, `ConsumerFactory`,
`StructureFactory`, `VehicleFactory`, `StandardLab`, `AdvancedLab`, `Observatory`,
`ReinforcedResidence`, `AdvancedResidence`, `CommonOreSmelter`, `Spaceport`,
`RareOreSmelter`, `GORF`, `Tokamak`.

**Vehicles** include: `CargoTruck`, `ConVec`, `Lynx`, `Panther`, `Tiger`, `Spider`,
`Scorpion`, `RoboSurveyor`, `RoboMiner`, `GeoCon`, `Scout`, `RoboDozer`, `EvacuationTransport`,
`RepairVehicle`, `Earthworker`, etc. (combat units `Lynx`/`Panther`/`Tiger`/`Spider`/
`Scorpion` and `GuardPost` accept a weapon via `CargoType`).

**Walls/tubes:** `Tube`, `Wall`, `LavaWall`, `MicrobeWall`.

**Beacon kinds (`MapID`):** `MiningBeacon`, `Fumarole`, `MagmaVent`.

> The full `map_id` enum is defined in the Outpost 2 SDK / `DotNetMissionSDK`; the lists
> above are the values OP2MissionEditor places and OP2OpmTools recognizes. Eden and
> Plymouth have different available structure sets (e.g. `MHDGenerator`, `Forum`,
> `ArachnidFactory`, `ReinforcedResidence` are Plymouth; `Tokamak`, `ConsumerFactory`,
> `Observatory` are Eden).

---

## 11. Minimal valid example

A complete 2‑player colony mission with empty starting bases (from `cnewmap.opm`):

```json
{
  "Disasters": [],
  "LevelDetails": {
    "LevelDescription": "No description",
    "MapName": "newworld.map",
    "MaxTechLevel": 12,
    "MissionType": "Colony",
    "NumPlayers": 2,
    "TechTreeName": "MULTITEK.TXT",
    "UnitOnlyMission": false
  },
  "MasterVariant": {
    "AutoLayouts": [],
    "Name": null,
    "Players": [
      {
        "Allies": [],
        "BotType": "None",
        "Color": "Blue",
        "DifficultyResources": [],
        "ID": 0,
        "IsEden": true,
        "IsHuman": true,
        "Resources": {
          "CenterView": { "X": 0, "Y": 0 },
          "CommonOre": 0,
          "CompletedResearch": [],
          "Food": 1000,
          "FreeMorale": false,
          "Kids": 10,
          "MoraleLevel": "Good",
          "RareOre": 0,
          "Scientists": 8,
          "SolarSatellites": 0,
          "TechLevel": 0,
          "Triggers": [],
          "Units": [],
          "WallTubes": [],
          "Workers": 14
        }
      }
    ],
    "TethysDifficulties": [],
    "TethysGame": {
      "Beacons": [],
      "DaylightEverywhere": false,
      "DaylightMoves": true,
      "InitialLightLevel": 0,
      "Markers": [],
      "MusicPlayList": { "RepeatStartIndex": 0, "SongIDs": [0] },
      "Triggers": [],
      "Wreckage": []
    }
  },
  "MissionVariants": [],
  "Regions": [],
  "SDKVersion": "0",
  "Triggers": []
}
```

---

## 12. Quick field index

| Path | Type |
|---|---|
| `SDKVersion` | string |
| `LevelDetails.{LevelDescription, MapName, TechTreeName}` | string |
| `LevelDetails.{MaxTechLevel, NumPlayers}` | int |
| `LevelDetails.MissionType` | enum string |
| `LevelDetails.UnitOnlyMission` | bool |
| `*Variant.Name` | string \| null |
| `*Variant.Players[]` | Player[] |
| `*Variant.{TethysGame, TethysDifficulties[]}` | GameData |
| `*Variant.AutoLayouts[]` | AutoLayout[] |
| `Player.{ID}` | int |
| `Player.{IsEden, IsHuman}` | bool |
| `Player.{Color, BotType}` | enum string |
| `Player.AIImpl` | enum string (optional) |
| `Player.Allies[]` | int[] |
| `Player.Resources`, `Player.DifficultyResources[]` | Resources |
| `Resources.{CommonOre, RareOre, Food, Kids, Workers, Scientists, SolarSatellites, TechLevel}` | int |
| `Resources.CompletedResearch[]` | int[] |
| `Resources.MoraleLevel` | enum string |
| `Resources.FreeMorale` | bool |
| `Resources.CenterView` | Point |
| `Resources.Units[]` | Unit[] |
| `Resources.WallTubes[]` | WallTube[] |
| `GameData.{DaylightEverywhere, DaylightMoves}` | bool |
| `GameData.InitialLightLevel` | int |
| `GameData.MusicPlayList` | {RepeatStartIndex:int, SongIDs:int[]} |
| `GameData.Beacons[]` | Beacon[] |
| `GameData.Markers[]` | Marker[] |
| `GameData.Wreckage[]` | Wreckage[] |
| `GameData.WallTubes[]` | WallTube[] (optional) |
| `Unit.{ID, CargoAmount, MinDistance, SpawnDistance}` | int |
| `Unit.MaxTubes` | int \| null |
| `Unit.{TypeID, Direction, CargoType, BarVariant, BarYield}` | enum string |
| `Unit.Health` | float (0–1) |
| `Unit.{Lights, CreateWall, IgnoreLayout}` | bool |
| `Unit.Position`, `*.CenterView`, `*.Position` | Point |
| `Beacon.{ID}` | int |
| `Beacon.{MapID, OreType, BarYield, BarVariant}` | enum string |
| `Marker.{ID}` | int (optional), `MarkerType` enum string |
| `Wreckage.{ID}` | int (optional), `TechID` **int**, `IsVisible` bool |
| `WallTube.TypeID` | enum string |
| `WallTube.Position` | Point (canonical key; some builds emit `Location`) |
| `Point.{X, Y}` (`DataLocation`) | int |
| `DataRect.{MinX, MinY, MaxX, MaxY}` | int |
| `AutoLayout.{PlayerID}` | int; `BaseCenterPt` Point; `Units[]` Unit[] |
| `MissionRoot.Disasters[]` | Disaster[] (`DisasterData`) |
| `MissionRoot.Regions[]` | Region[] (`RegionData`: Name, Rect) |
| `MissionRoot.Triggers[]` | **OP2TriggerData[]** (flat) |
| `GameData.Triggers[]`, `Resources.Triggers[]` | **TriggerData[]** (condition/action) |
| `Disaster.{ID, Size, Duration, StartTime, EndTime, MinDelay, MaxDelay}` | int |
| `Disaster.Type` | enum string; `SrcRect`/`DestRect` DataRect |
| `TriggerData.{ID}` int, `{Enabled}` bool, `{EventType}` enum, `Condition[]`, `Actions[]` | |
| `TriggerConditionData.{Type, Subject, Comparison, Value, Value2, Value3}` | string(s) |
| `TriggerActionData.{Type, Modifier, Subject*, Value..Value5}` | string(s), `SubjectQuantity` int |
| `OP2TriggerData` | flat fields — see [§9.3](#93-op2triggerdata-missionroottriggers) |

---

*Compiled from the OP2DotNetMissionSDK source (`DotNetMissionReader`,
`MissionReader/Json/`) as the authoritative model, real `.opm` samples, and OP2OpmTools
(`Unit.vb`, `fMain.vb`). JSON keys come from each member's `[DataMember(Name=…)]`; enum
values are serialized by name. Disasters / Triggers / Regions are defined by the SDK but
left empty by OP2MissionEditor (no authoring UI).*
