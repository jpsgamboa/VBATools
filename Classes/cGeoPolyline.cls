VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGeoPolyline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCoordinates As Collection
Private pDistance As Double
Private pTravelTime As Double
Private pToll As Boolean

Public Sub init()
    Set pCoordinates = New Collection
End Sub

Public Function addCoordinate(coord As cGeoCoordinate)
    pCoordinates.add (coord)
End Function

Public Function getCoordinate(index As Integer) As cGeoCoordinate
    getCoordinate = pCoordinates(index)
End Function

Public Property Get coordinates() As Collection: Set coordinates = pCoordinates: End Property

Public Property Get distance() As Double: distance = pDistance: End Property
Public Property Let distance(value As Double): pDistance = value: End Property

Public Property Get travelTime() As Double: travelTime = pTravelTime: End Property
Public Property Let travelTime(value As Double): pTravelTime = value: End Property

Public Property Get HasToll() As Boolean: HasToll = pToll: End Property
Public Property Let SetToll(value As Boolean): pToll = value: End Property

