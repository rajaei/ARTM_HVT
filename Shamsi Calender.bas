Attribute VB_Name = "Calender"
Option Explicit

Private Month_Name, Spring_Fall
Private Time_Difference, Time_Client
Private Base_Year


Private Sub Get_Date(ByVal Days, Sal, Mah, Rooz)
   Dim Years, Year_Length
   Do While Days >= 0
     If Kabiseh(Years) Then
        Year_Length = 366
     Else
        Year_Length = 365
     End If
     If Days - Year_Length >= 0 Then
        Years = Years + 1
        Days = Days - Year_Length
     Else
        Sal = Base_Year + Years
        If Days <= 185 Then
           Mah = 1 + (Days \ 31)
           Rooz = 1 + (Days Mod 31)
        Else
           Days = Days - 186
           Mah = 7 + (Days \ 30)
           Rooz = 1 + (Days Mod 30)
        End If
        Exit Sub
     End If
   Loop
End Sub
Private Function Kabiseh(ByVal Years)
   Dim Temp
   Kabiseh = False
   Temp = (Base_Year + Years) - 1309
   If (((Temp Mod 32) - (Temp \ 32)) Mod 4) = 0 Then Kabiseh = True
End Function
Public Property Let SFhour(x)
   Spring_Fall = x
End Property
Public Property Let Time_Diff(ByVal t)
   Time_Difference = t
End Property
Public Property Let state(ByVal S)
   Month_Name = S
End Property
Public Function To_HejriChar(ByVal what_date, Optional Month_Name)
   Dim Days, Day_Name, Day_Number, Temp_Days, Months
   Spring_Fall = False
   If IsMissing(Month_Name) Then Month_Name = 0

   Time_Difference = #12:00:00 AM#
   Base_Year = 1332

   Months = Array("ÝÑæÑÏíä", "ÇÑÏíÈåÔÊ", "ÎÑÏÇÏ", "ÊíÑ", "ãÑÏÇÏ", "ÔåÑíæÑ", "ãåÑ", "ÂÈÇä", "ÂÐÑ", "Ïí", "Èåãä", "ÇÓÝäÏ")

   Day_Name = Array("í˜ÔäÈå", "ÏæÔäÈå", "Óå ÔäÈå", "åÇÑÔäÈå", "äÌÔäÈå", "ÌãÚå", "ÔäÈå")
   Days = DateDiff("d", #3/21/1953#, what_date)
   Day_Number = Weekday(what_date)
   Dim Year_Length, Sal, Mah, Rooz, temp_date
   If FormatDateTime(what_date + Time_Difference, vbShortDate) <> FormatDateTime(what_date, vbShortDate) Then
      Days = Days + 1
      Day_Number = (Day_Number + 1)
      If Day_Number = 8 Then Day_Number = 1
   End If
   Time_Client = FormatDateTime(what_date + Time_Difference, vbLongTime)
   Call Get_Date(Days, Sal, Mah, Rooz)
   If ((Mah >= 1 And Mah <= 6) And Not ((Mah = 1 And Rooz = 1) Or (Mah = 6 And Rooz = 31))) And Spring_Fall = True Then
      If FormatDateTime(what_date + Time_Difference + #1:00:00 AM#, vbShortDate) <> FormatDateTime(what_date + Time_Difference, vbShortDate) Then
         Temp_Days = Days + 1
         Day_Number = (Day_Number + 1)
         If Day_Number = 8 Then Day_Number = 1
      Else
         Temp_Days = Days
      End If
      Time_Client = FormatDateTime(what_date + Time_Difference + #1:00:00 AM#, vbLongTime)
      If Temp_Days <> Days Then
         Days = Temp_Days
         If Rooz = 30 And Mah = 6 Then
            If DateDiff("n", Time_Client, #1:00:00 AM#) <= 60 And DateDiff("n", Time_Client, #1:00:00 AM#) >= 0 Then
               Time_Client = FormatDateTime(what_date + Time_Difference, vbLongTime)
               Days = Days - 1
               If Day_Number = 1 Then
                  Day_Number = 7
               Else
                  Day_Number = Day_Number - 1
               End If
            End If
         End If
         Call Get_Date(Days, Sal, Mah, Rooz)
      End If
   End If
   If Month_Name = 0 Then
      If Rooz < 10 Then Rooz = "0" & Rooz
      If Mah < 10 Then Mah = "0" & Mah
      To_HejriChar = Sal & "/" & Mah & "/" & Rooz
   ElseIf Month_Name = 1 Then
      To_HejriChar = Rooz & " " & Months(Mah - 1) & " " & Sal
   ElseIf Month_Name = 2 Then
      To_HejriChar = Day_Name(Day_Number - 1) & " " & Sal & "/" & Mah & "/" & Rooz
   ElseIf Month_Name = 3 Then
      To_HejriChar = Day_Name(Day_Number - 1) & " " & Rooz & " " & Months(Mah - 1) & " " & Sal
   End If
End Function

Private Sub Class_Initialize()
   Spring_Fall = False
   Month_Name = 0
   Time_Difference = #12:00:00 AM#
   Base_Year = 1332
End Sub




