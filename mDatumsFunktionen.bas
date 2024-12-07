Attribute VB_Name = "mDatumsFunktionen"
Option Explicit

Sub AddFunctionTooltip()
    Application.MacroOptions _
        Macro:="TagDesJahres", _
        Description:="Berechnet den Tag des Jahres für das angegebene Datum .Falls kein Datum angegeben ist, wird das aktuelle Datum verwendet", _
        ArgumentDescriptions:=Array("Datum as Date")
End Sub

Public Function getAktuellesDatum() As String
    Dim aktdatum As String
    
    aktdatum = "Aktuells Datum: " & CStr(Date) & _
    Space(1) & " | Wochentag: " & getWeekday(Date) & _
    Space(1) & " | Tag des Monats: " & getTagDesMonatsZweistellig(Date) & _
    Space(1) & " | Monat: " & getMonatZweistellig(Date) & _
    Space(1) & " | Jahr: " & CStr(Year(Date)) & _
    Space(1) & " | Tag des Jahres: " & CStr(TagDesJahres) & _
    Space(1) & " | Kalenderwoche: " + Kalenderwoche(Date)
    
    getAktuellesDatum = aktdatum
End Function

Public Function TagDesJahres(Optional datum As Date) As Integer
Attribute TagDesJahres.VB_Description = "Berechnet den Tag des Jahres für das angegebene Datum .Falls kein Datum angegeben ist, wird das aktuelle Datum verwendet"
Attribute TagDesJahres.VB_ProcData.VB_Invoke_Func = " \n14"
    'Berechnet den Tag des Jahres für das angegebene Datum .Falls kein Datum angegeben ist, wird das aktuelle Datum verwendet
    If datum = 0 Then datum = Date
    TagDesJahres = datum - DateSerial(Year(Date), 1, 1) + 1
End Function

Public Function DatumsDifferenz_in_Monate(Erstes_Datum As Date, Zweites_Datum As Date) As Long
   DatumsDifferenz_in_Monate = DateDiff("m", Erstes_Datum, Zweites_Datum)
End Function

Public Function DatumsDifferenz_in_Tage(Erstes_Datum As Date, Zweites_Datum As Date) As Long
   DatumsDifferenz_in_Tage = DateDiff("d", Erstes_Datum, Zweites_Datum)
End Function

Public Function DatumsDifferenz_in_Jahre(Erstes_Datum As Date, Zweites_Datum As Date) As Long
   DatumsDifferenz_in_Jahre = DateDiff("yyyy", Erstes_Datum, Zweites_Datum)
End Function

Public Function Kalenderwoche(d As Date) As String
    Dim t
    t = DateSerial(Year(d + (8 - Weekday(d)) Mod 7 - 3), 1, 1)
    Kalenderwoche = ((d - t - 3 + (Weekday(t) + 1) Mod 7)) \ 7 + 1
End Function

Public Function getWeekday(ByVal datum As Date) As String
    Dim wt As Integer
    wt = Weekday(datum)
    
    Select Case wt
        Case 1
          getWeekday = "Sonntag"
        Case 2
          getWeekday = "Montag"
        Case 3
          getWeekday = "Dienstag"
        Case 4
          getWeekday = "Mittwoch"
        Case 5
          getWeekday = "Donnerstag"
        Case 6
          getWeekday = "Freitag"
        Case 7
          getWeekday = "Samstag"
    End Select

End Function

Public Function getTagDesMonatsZweistellig(d As Date) As String
    Dim tag As String
    tag = Day(Date)
    If Len(tag) = 1 Then
        tag = "0" & tag
    End If
    getTagDesMonatsZweistellig = tag
End Function

Public Function getMonatZweistellig(d As Date) As String
    Dim monat As String
    monat = Month(d)
    If Len(monat) = 1 Then
        monat = "0" & monat
    End If
    getMonatZweistellig = monat
End Function
