 'From c code posted here: http://www.psa.es/sdg/sunpos.htm
 'VB.NET Conversion posted here: http://www.vbforums.com/showthread.php?832645-Solar-position-calculator
 'converted for HomeSeer use by Sparkman v1.0

' old lat/long 38.9894 -77.1370

Imports System.Math

Sub Main(ByVal Parms As String)
    Dim Debug As Boolean = False
    Dim logName As String = "Solar Position"
    Dim dLatitude As Double = 38.998642
    Dim dLongitude As Double = -77.126819
    Dim hsAzimuthDevice As Integer = 2546
    Dim hsAltitudeDevice As Integer = 2545

    Dim pi As Double = 3.14159265358979
    Dim rad As Double = pi / 180
    Dim dEarthMeanRadius As Double = 6371.01
    Dim dAstronomicalUnit As Double = 149597890

    Dim iYear As Integer = DateTime.UtcNow.Year
    Dim iMonth As Integer = DateTime.UtcNow.Month
    Dim iDay As Integer = DateTime.UtcNow.Day
    Dim dHours As Double = DateTime.UtcNow.Hour
    Dim dMinutes As Double = DateTime.UtcNow.Minute
    Dim dSeconds As Double = DateTime.UtcNow.Second

    Dim dZenithAngle As Double
    Dim dZenithAngleParallax As Double
    Dim dAzimuth As Double
    Dim dAltitudeAngle As Double
    Dim dElapsedJulianDays As Double
    Dim dDecimalHours As Double
    Dim dEclipticLongitude As Double
    Dim dEclipticObliquity As Double
    Dim dRightAscension As Double
    Dim dDeclination As Double
    Dim dY As Double
    Dim dX As Double
    Dim dJulianDate As Double
    Dim liAux1 As Integer
    Dim liAux2 As Integer
    Dim dMeanLongitude As Double
    Dim dMeanAnomaly As Double
    Dim dOmega As Double
    Dim dSin_EclipticLongitude As Double
    Dim dGreenwichMeanSiderealTime As Double
    Dim dLocalMeanSiderealTime As Double
    Dim dLatitudeInRadians As Double
    Dim dHourAngle As Double
    Dim dCos_Latitude As Double
    Dim dSin_Latitude As Double
    Dim dCos_HourAngle As Double
    Dim dParallax As Double

    Try

        ' Calculate difference in days between the current Julian Day and JD 2451545.0, which is noon 1 January 2000 Universal Time
        ' Calculate time of the day in UT decimal hours
        dDecimalHours = dHours + (dMinutes + dSeconds / 60.0) / 60.0
        ' Calculate current Julian Day
        liAux1 = (iMonth - 14) \ 12
        liAux2 = (1461 * (iYear + 4800 + liAux1)) \ 4 + (367 * (iMonth - 2 - 12 * liAux1)) \ 12 - (3 * ((iYear + 4900 + liAux1) \ 100)) \ 4 + iDay - 32075
        dJulianDate = CDbl(liAux2) - 0.5 + dDecimalHours / 24.0
        ' Calculate difference between current Julian Day and JD 2451545.0
        dElapsedJulianDays = dJulianDate - 2451545.0
        If Debug Then hs.writelog(logName,"Elapsed Julian Days Since 2000/01/01: " & CStr(dElapsedJulianDays))

        ' Calculate ecliptic coordinates (ecliptic longitude and obliquity of the ecliptic in radians but without limiting the angle to be less than 2*Pi
        ' (i.e., the result may be greater than 2*Pi)
        dOmega = 2.1429 - 0.0010394594 * dElapsedJulianDays
        dMeanLongitude = 4.895063 + 0.017202791698 * dElapsedJulianDays ' Radians
        dMeanAnomaly = 6.24006 + 0.0172019699 * dElapsedJulianDays
        dEclipticLongitude = dMeanLongitude + 0.03341607 * Math.Sin(dMeanAnomaly) + 0.00034894 * Math.Sin(2 * dMeanAnomaly) - 0.0001134 - 0.0000203 * Math.Sin(dOmega)
        dEclipticObliquity = 0.4090928 - 0.000000006214 * dElapsedJulianDays + 0.0000396 * Math.Cos(dOmega)
        If Debug Then hs.writelog(logName,"Ecliptic Longitude: " & CStr(dEclipticLongitude))
        If Debug Then hs.writelog(logName,"Ecliptic Obliquity: " & CStr(dEclipticObliquity))

        ' Calculate celestial coordinates ( right ascension and declination ) in radians but without limiting the angle to be less than 2*Pi (i.e., the result may be greater than 2*Pi)
        dSin_EclipticLongitude = Math.Sin(dEclipticLongitude)
        dY = Math.Cos(dEclipticObliquity) * dSin_EclipticLongitude
        dX = Math.Cos(dEclipticLongitude)
        dRightAscension = Math.Atan2(dY, dX)
        If dRightAscension < 0.0 Then
            dRightAscension = dRightAscension + (2 * pi)
        End If
        dDeclination = Math.Asin(Math.Sin(dEclipticObliquity) * dSin_EclipticLongitude)
        If Debug Then hs.writelog(logName,"Declination: " & CStr(dDeclination))

        ' Calculate local coordinates ( azimuth and zenith angle ) in degrees
        dGreenwichMeanSiderealTime = 6.6974243242 + 0.0657098283 * dElapsedJulianDays + dDecimalHours
        dLocalMeanSiderealTime = (dGreenwichMeanSiderealTime * 15 + dLongitude) * rad
        dHourAngle = dLocalMeanSiderealTime - dRightAscension
        If Debug Then hs.writelog(logName,"Hour Angle: " & CStr(dHourAngle))
        dLatitudeInRadians = dLatitude * rad
        dCos_Latitude = Math.Cos(dLatitudeInRadians)
        dSin_Latitude = Math.Sin(dLatitudeInRadians)
        dCos_HourAngle = Math.Cos(dHourAngle)
        dZenithAngle = (Math.Acos(dCos_Latitude * dCos_HourAngle * Math.Cos(dDeclination) + Math.Sin(dDeclination) * dSin_Latitude))
        dY = -Math.Sin(dHourAngle)
        dX = Math.Tan(dDeclination) * dCos_Latitude - dSin_Latitude * dCos_HourAngle
        dAzimuth = Math.Atan2(dY, dX)
        If dAzimuth < 0.0 Then
            dAzimuth = dAzimuth + (2 * pi)
        End If
        dAzimuth = dAzimuth / rad

        If Debug Then hs.writelog(logName,"Azimuth: " & CStr(dAzimuth))
        hs.setdevicevaluebyref(hsAzimuthDevice,dAzimuth,True)

        ' Parallax Correction
        dParallax = (dEarthMeanRadius / dAstronomicalUnit) * Math.Sin(dZenithAngle)
        dZenithAngleParallax = (dZenithAngle + dParallax) / rad
        dAltitudeAngle = (dZenithAngleParallax * -1) + 90

        If Debug Then hs.writelog(logName,"Altitude Angle: " & CStr(dAltitudeAngle))
        hs.setdevicevaluebyref(hsAltitudeDevice,dAltitudeAngle,True)

    Catch ex As Exception
        hs.WriteLog(logName, "Exception " & ex.ToString)
    End Try

End Sub
