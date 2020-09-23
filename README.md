<div align="center">

## Determine Offset in minutes from GMT


</div>

### Description

Determines the number of minutes offset from GMT using WMI to determine daylight and timezone bias.
 
### More Info
 
Using Win95 Serv 2 or greater, 98, ME, NT, 2000

Minutes of bias off of GMT


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Karl D\. Parker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/karl-d-parker.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/karl-d-parker-determine-offset-in-minutes-from-gmt__1-26971/archive/master.zip)

### API Declarations

Uses WMI interfaces


### Source Code

```
Public Function GetGMTOffSet() As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Returns the GMT to Local time Offset. The return is provided in Minutes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  'Create and set a new object to the Windows Management Instrument
  Dim objWinManInstrument As Object
  Set objWinManInstrument = GetObject("WinMgmts:")
  'Create and set object to an instance of the SystemTimeZone object withing the Windows Management Instrument
  Dim objSysTimeZoneInfo As Object
  Set objSysTimeZoneInfo = objWinManInstrument.InstancesOf("win32_SystemTimeZone")
  'Create an object collection that we need to extract informaiton from the objSysTimeZoneInfo object
  Dim objObjectCollection As Object
  'This object will actually hold the Time Zone information extracted from the objSysTimeZoneInfo object
  Dim objTimeZone As Object
  'Holds the difference in minutes between GMT time and local time as defined by the systems setting of your computer
  Dim intMinMin As Integer
  'Holds the start date of day light savings time as defined by the system setting of your computer
  Dim dtmDayLightDateTime As Date
  'Holds the return to standard time as defined by the system setting of your computer
  Dim dtmStandardDateTime As Date
  'Holds the local time converted to GMT date/time
  Dim dtmGMTDateTime As Date
  'This is the only way I could get the actual Time Zone information from the objSysTimeZoneInfo object.
  'There is only one object in the objSysTimeZoneInfo object.
  For Each objObjectCollection In objSysTimeZoneInfo
    'Sets the objTimeZone to the time zone informaiton
    Set objTimeZone = objWinManInstrument.Get(objObjectCollection.Setting)
    'I care about the first object.
    Exit For
  Next
  'Some places in the world don't have a daylight savings time. For those places we need only look at the
  'time zone bias. If a daylight savings time bias exist we do the if poriton of the statement. If now
  'daylight savings time bias exists we do the else portin.
  If objTimeZone.DaylightBias <> 0 Then
    'Create the date/time for beginning of day light savings time as defined by your computer system setting and held in the
    'objTimeZone object
    dtmDayLightDateTime = CDate(Format(objTimeZone.DaylightMonth & "/" & objTimeZone.DaylightDay & "/" & Format(Date, "yyyy"), _
      "Short Date") & " " & Format(objTimeZone.DaylightHour & ":00", "Short Time"))
    'Create the date/time for return to standard time as defined by your computer system setting and held in the
    'objTimeZone object
    dtmStandardDateTime = CDate(Format(objTimeZone.StandardMonth & "/" & objTimeZone.StandardDay & "/" & Format(Date, "yyyy"), _
      "Short Date") & " " & Format(objTimeZone.StandardHour & ":00", "Short Time"))
    'Check to see if we are in daylight savings time or standard time.
    If Now > dtmDayLightDateTime And Now < dtmStandardDateTime Then
      'If daylight savings time we need to set the intMinMin diferential to include the time zone bias
      'and the daylight savings time bias
      intMinMin = objTimeZone.Bias + objTimeZone.DaylightBias
    Else
      'If we are in standard time we set the intMinMin diferential to include only the time zone bias
      intMinMin = objTimeZone.Bias
    End If
  Else
    'No dalight savings time bias exists. Lood at only the time zone bias
    intMinMin = objTimeZone.Bias
  End If
  'SEt the function return to the calculated GMT Date/Time
  GetGMTOffSet = intMinMin
End Function
```

