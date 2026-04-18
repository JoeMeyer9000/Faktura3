VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *****************************************************
' Klasse JApplication
' *****************************************************

' Wenn man über die globale Instanz des APplication Objekts eine Excel Datei öffnet oder
' erzeugt, wird diese auch angezeigt. Das ist ok, wenn man einen neuen Beleg anlegt, aber
' wenn man z. Bsp. druckt oder Belege checkt oder Stornos erzeugt, führt dies zum
' kurzzeitigen Aufflackern des Belegs, obwohl der Benutzer diesen gar nicht sehen will.
' In solchen Fällen erzeugt man dann eine Instanz von JApplication und benutzt diese
' Instanz, um ein Workbook zu erzeugen oder zu öffnen. Da JApplication mit Visible = False
' erzeugt wird, sind die Workbooks von JApplication nicht sichtbar.
' Die Instanz wird mit den Constructor Class_Initialize() angelegt und in Class_Terminate()
' auch wieder entfernt. Das bedeutet aber auch, dass ein Beleg, der über diese Instanz
' geöffnet wird, auch wieder verschwindet, wenn die Instanz entfernt wird. Das passiert nicht
' mit der globalen Instanz Application.Workbooks, die immer geöffnet bleibt

' Anwendung:
' Dim app as new JApplication

' Beim ersten Aufruf von app wird die unsichtbare Instanz einer Application erzeugt
' Zugriff auf das unsichtbare Application Objeckt erhält man dann mit app.Instance
' Ergo: statt Application.Worksbooks.New schreibt man app.Instance.Worksbooks.New

Private m_InvisibleApp As Application

' =====================================================
' Public
' =====================================================

' Liefert die Application Instanz, die beim ersten Aufruf instanziiert wird
Public Property Get Instance() As Application
    If m_InvisibleApp Is Nothing Then
        Set m_InvisibleApp = New Application
        m_InvisibleApp.visible = False
        m_InvisibleApp.DisplayAlerts = False
        m_InvisibleApp.ScreenUpdating = False
        m_InvisibleApp.DisplayStatusBar = False
    End If
    Set Instance = m_InvisibleApp
End Property

' =====================================================
' Private
' =====================================================

' Destruktor. Gibt die Application Instanz wieder frei
Private Sub Class_Terminate()
    If Not (m_InvisibleApp Is Nothing) Then
        m_InvisibleApp.Quit
        Set m_InvisibleApp = Nothing
    End If
End Sub
