Attribute VB_Name = "mUmfangFlaechenBerechnung"
Option Explicit

'Konstante pi deklarieren
Public Const pi As Double = 3.14159265358979


'Das Rechteck
'Ein Rechteck ist ein Viereck, dessen Winkel 90° betragen. Dadurch sind gegenüberliegende
'Seiten gleich groß. Seine Diagonalen sind ebenfalls gleich lang.
'Formel für den Umfang: U = a + b + a + b = 2a+2b
Public Function BerechneUmfangRechteck(a_cm As Double, b_cm As Double) As Double
        Dim umfang As Double
    umfang = (2 * a_cm) + (2 * b_cm)
    BerechneUmfangRechteck = Round(umfang)
End Function

'Der Flächeninhalt („Länge mal Breite“) ist ebenso einfach – man multipliziert die Seitenlängen mit einander:
' A= a + b
Public Function BerechneFlaecheninhaltRechteck(a_cm As Double, b_cm As Double) As Double
    Dim flaecheninhalt As Double
    flaecheninhalt = a_cm * b_cm
    BerechneFlaecheninhaltRechteck = Round(flaecheninhalt)
End Function


'Umfang Kreis berechnen
'Der Umfang u eines Kreises mit dem Radius r berechnet sich mit der Formel:
'Kreisradius r
'Durchmesser d = 2 * r
'Umfang Kreis = 2 * pi * r
'Fläche Kreis = A = pi * r2
Public Function BerechneKreisumfang(d_cm As Double) As Double
    Dim kreisumfang As Double
    kreisumfang = 2 * pi * (d_cm / 2)
    BerechneKreisumfang = Round(kreisumfang, 2)
End Function

Public Function BerechneKreisflaeche(d_cm As Double) As Double
    Dim kreisflaeche As Double
    Dim radius_cm As Double
    radius_cm = d_cm / 2
    kreisflaeche = pi * (radius_cm ^ 2)
    BerechneKreisflaeche = Round(kreisflaeche, 3)
End Function


