Attribute VB_Name = "Geographic"
Option Explicit

Function DistanceCrowFlies(dlat1, dlon1, dlat2, dlon2) As Variant
    Dim PI As Variant, lat1 As Variant, lat2 As Variant, lon1 As Variant, lon2 As Variant, cosX As Variant
    
    PI = Application.PI()
    Const earthRadius = 6371

    lat1 = dlat1 * PI / 180
    lat2 = dlat2 * PI / 180
    lon1 = dlon1 * PI / 180
    lon2 = dlon2 * PI / 180

    cosX = sIn(lat1) * sIn(lat2) + Cos(lat1) _
      * Cos(lat2) * Cos(lon1 - lon2)
    DistanceCrowFlies = earthRadius * Application.Acos(cosX)
End Function

Function DistanceCrowFlies2(latlon1, latlon2) As Variant
    Dim PI As Variant, lat1 As Variant, lat2 As Variant, lon1 As Variant, lon2 As Variant, cosX As Variant
    Dim dlat1 As Variant, dlat2 As Variant, dlon1 As Variant, dlon2 As Variant
    dlat1 = Split(latlon1, ",")(0)
    dlat2 = Split(latlon2, ",")(0)
    dlon1 = Split(latlon1, ",")(1)
    dlon2 = Split(latlon2, ",")(1)

    DistanceCrowFlies2 = DistanceCrowFlies(dlat1, dlon1, dlat2, dlon2)
End Function




