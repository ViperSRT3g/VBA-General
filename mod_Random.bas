Attribute VB_Name = "mod_Random"
Option Explicit

Public Function Random(Lowerbound As Long, Upperbound As Long) As Long
    Randomize
    Random = Rnd * (Upperbound - Lowerbound) + Lowerbound
End Function
