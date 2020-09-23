VERSION 5.00
Begin VB.UserControl WeatherWatch 
   Alignable       =   -1  'True
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "WeatherWatch.ctx":0000
   ScaleHeight     =   465
   ScaleWidth      =   495
   ToolboxBitmap   =   "WeatherWatch.ctx":08CA
End
Attribute VB_Name = "WeatherWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private objHTTP     As New MSXML.XMLHTTPRequest

Private strCity     As String
Private strState    As String
Private strResponse As String

Private Type WeatherGrid
    CurrentStat     As String
    Temp            As String
    Wind            As String
    DewPoint        As String
    Humidity        As String
    Visibility      As String
    Barometer       As String
    Sunrise         As String
    Sunset          As String
End Type

Private objWeather As WeatherGrid
Private Function CheckCity(City As String)

Dim strCity      As String
Dim strTempCity  As String
Dim strChrHolder As String
Dim Counter      As Integer

strCity = Trim(City)

'If city name is more then two words we need to add an underscore to join them
If InStr(strCity, " ") > 0 Then
    For Counter = 1 To Len(strCity)
        strChrHolder = Mid(strCity, Counter, 1)
        If strChrHolder <> Chr(32) Then
            strTempCity = strTempCity & strChrHolder
        Else
            strTempCity = strTempCity & "_"
        End If
    Next
    strCity = strTempCity
End If

CheckCity = strCity

End Function

Public Sub Connect()

Dim strWebPage  As String
Dim strResponse As String

strWebPage = "http://www.weather.com/weather/cities/us_" & strState & "_" & strCity & ".html"
objHTTP.open "GET", strWebPage, False
objHTTP.send

If objHTTP.Status <> "200" Then
    MsgBox "Cannot find City/State combination, please " & vbNewLine & _
           "make sure you spelled the City correctly", vbCritical, "Weather Today"
    Exit Sub
End If

strResponse = objHTTP.responseText

objWeather = ParseData(strResponse)


End Sub

Public Property Get DewPoint() As String

DewPoint = objWeather.DewPoint & " deg"

End Property

Private Function ParseData(Request As String) As WeatherGrid

Dim strRequest     As String
Dim StartFrom      As Long
Dim EndAt          As Long
Dim RetVal         As Long
Dim strCurrentStat As String
Dim strTemp        As String
Dim strWind        As String
Dim strDewPoint    As String
Dim strHumitity    As String
Dim strVisibility  As String
Dim strBarometer   As String
Dim strSunrise     As String
Dim strSunset      As String

strRequest = Request

'Search Webpage for the Data we need

'Get Current Status
RetVal = InStr(strRequest, "as reported at")
StartFrom = RetVal
RetVal = InStr(StartFrom, strRequest, "<B>")
StartFrom = RetVal + 3
EndAt = InStr(StartFrom, strRequest, "</B>")
strCurrentStat = Mid(strRequest, StartFrom, EndAt - StartFrom)

'Get Current Temp
RetVal = InStr(strRequest, "Temp:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "&deg;F")
    StartFrom = RetVal
    strTemp = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Wind Speed
RetVal = InStr(StartFrom, strRequest, "Wind:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "mph")
    StartFrom = RetVal
    strWind = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current DewPoint
RetVal = InStr(StartFrom, strRequest, "Dewpoint:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "&deg;F")
    StartFrom = RetVal
    strDewPoint = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Humidity
RetVal = InStr(StartFrom, strRequest, "Rel. Humidity:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "%")
    StartFrom = RetVal
    strHumidity = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Visibility
RetVal = InStr(StartFrom, strRequest, "Visibility:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "miles")
    StartFrom = RetVal
    strVisibility = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Barometer
RetVal = InStr(StartFrom, strRequest, "Barometer:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "inches")
    StartFrom = RetVal
    strBarometer = CheckOther(Mid(strRequest, StartFrom - 6, 6))
End If

'Get Current Sunrise
RetVal = InStr(StartFrom, strRequest, "Sunrise:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "am")
    StartFrom = RetVal
    strSunrise = CheckOther(Mid(strRequest, StartFrom - 6, 6))
End If

'Get Current Sunset
RetVal = InStr(StartFrom, strRequest, "Sunset:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "pm")
    StartFrom = RetVal
    strSunset = CheckOther(Mid(strRequest, StartFrom - 6, 6))
End If

ParseData.CurrentStat = strCurrentStat
ParseData.Temp = strTemp
ParseData.Wind = strWind
ParseData.DewPoint = strDewPoint
ParseData.Humidity = strHumidity
ParseData.Visibility = strVisibility
ParseData.Barometer = strBarometer
ParseData.Sunrise = strSunrise
ParseData.Sunset = strSunset

End Function

Private Function CheckIt(Tmp As String) As String

Dim strTmp       As String
Dim strTempTmp   As String
Dim strChrHolder As String
Dim Counter      As Integer

strTmp = Tmp

For Counter = 1 To Len(strTmp)
    strChrHolder = Mid(strTmp, Counter, 1)
    If IsNumeric(strChrHolder) Then
        strTempTmp = strTempTmp & strChrHolder
    End If
Next

CheckIt = strTempTmp

End Function

Private Function CheckOther(Tmp As String) As String

Dim strTmp       As String
Dim strTempTmp   As String
Dim strChrHolder As String
Dim Counter      As Integer

strTmp = Tmp

For Counter = 1 To Len(strTmp)
    strChrHolder = Mid(strTmp, Counter, 1)
    If strChrHolder <> ">" Then
        strTempTmp = strTempTmp & strChrHolder
    End If
Next

CheckOther = strTempTmp

End Function

Public Property Get City() As String

City = strCity

End Property

Public Property Let City(ByVal strNewValue As String)

strCity = LCase(CheckCity(strNewValue))

End Property

Public Property Get State() As String

State = strState

End Property

Public Property Let State(ByVal strNewValue As String)

strState = LCase(strNewValue)

End Property

Public Property Get CurrentStatus() As String

CurrentStatus = objWeather.CurrentStat

End Property



Public Property Get Temperature() As String

Temperature = objWeather.Temp & " deg"

End Property


Public Property Get WindSpeed() As String

WindSpeed = objWeather.Wind & " mph"

End Property





Public Property Get Humidity() As String

Humidity = objWeather.Humidity & " %"

End Property


Public Property Get Visibility() As String

Visibility = objWeather.Visibility & " miles"

End Property


Public Property Get Barometer() As String

Barometer = objWeather.Barometer & " inches"

End Property


Public Property Get Sunrise() As String

Sunrise = objWeather.Sunrise & " am"

End Property



Public Property Get Sunset() As String

Sunset = objWeather.Sunset & " pm"

End Property

