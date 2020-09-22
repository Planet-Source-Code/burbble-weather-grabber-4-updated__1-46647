Attribute VB_Name = "modData"
Public CurrentData As String
Public RadarData As String
Public AnimateData As String
Public ForecastData As String
Public AlertData As String

Public AlertURL() As String
Public AlertCaption() As String
Public AlertCount As Integer
Public AlertPreviousCount As Integer
Public AlertPlaySound As Boolean

Public ZipCode As String
Public City As String
Public RefreshRate As Integer

Public JustCanceled As Boolean
Public ChangedZip As Boolean
Public CheckNow As Boolean
Public HideMe As Boolean

'Registry Settings
Public RegZip As String
Public RegFirst As String
Public RegRefresh As String
Public RegBackground As String
Public RegLoad As String
Public RegEverRefresh As String
Public RegInternet As String
Public RegNoCache As String
Public RegFontColor As String
Public RegLineColor As String
Public RegIcon As String
Public RegConnect As String
Public RegSound As String
Public RegUseProxy As String
Public RegProxy As String
Public RegProxyUser As String
Public RegProxyPassword As String
Public RegProxyPort As String
Public RegAlertWindow As String
