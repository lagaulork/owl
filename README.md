' セルにデータを格納
Cells(i, 1).Formula = "=IF(RC[1]=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",IF(COUNTIF(RC[1]," & Chr(34) & "z_*" & Chr(34) & ")," & Chr(34) & "○" & Chr(34) & "," & Chr(34) & "-" & Chr(34) & "))"

Cells(i, 1).Formula = "=IF(RC[1]=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",IF(NOT(COUNTIF(RC[1]," & Chr(34) & "z_*" & Chr(34) & "))," & Chr(34) & "○" & Chr(34) & "," & Chr(34) & "-" & Chr(34) & "))"
            
select replace(Sort, 1, '-') from HELI_IT_Doc
