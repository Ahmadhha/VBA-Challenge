'Stock_Analyzer_PRO_V1.0
' Assumptions required for correct output:
' (1) No data entry erros, eg. unqual last rows from the input columns
' (2) The timestamps are always in chronological order, ie. sorted from the past to the future


' Compute the Stock Analyzer code for all sheets

Sub worksheet_loop()
    Dim xSh As Worksheet
    
    For Each xSh In Worksheets
        xSh.Select
        Call Clear_Previous
    Next
    
    MsgBox "Result Area Clear!" & vbNewLine & "Click OK to Continue Analysis..."
    
    For Each xSh In Worksheets
        xSh.Select
        Call VBA_Challenge
    Next
    
End Sub

' Code to clear result space in case of code re-runs

Sub Clear_Previous()

'Find and clear result space (for code re-runs)

Dim res_lrow As Long

res_lrow = Cells(Rows.Count, 9).End(xlUp).Row

Range(Cells(1, 9), Cells(res_lrow, 25)).Clear

End Sub


' Stock Analyzer code

Sub VBA_Challenge()


'Variable definitions

Dim i As Long
Dim dat_lrow As Long
Dim res_lrow As Long
Dim res_trow As Long
Dim tckr_row As Long

Dim t_vol As Double
Dim opn As Double
Dim cls As Double
Dim chg As Double
Dim max As Double
Dim min As Double
Dim m_vol As Double

Dim tckr As String

'Create results headers

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Range("o2").Value = "Greatest % Increase"
Range("o3").Value = "Greatest % Decrease"
Range("o4").Value = "Greatest Total Volume"

'Find input data last row

dat_lrow = Cells(Rows.Count, 1).End(xlUp).Row

Application.ScreenUpdating = False

' Summarize ticker names to results space and compute required info

res_trow = 2
tckr_row = 1

t_vol = 0
max = 0
min = 0
m_vol = 0

For i = 2 To dat_lrow
    
    If tckr_row = 1 Then
        
        opn = Range("C" & i).Value
    
    End If
    
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        tckr = Cells(i, 1).Value
        t_vol = t_vol + Cells(i, 7).Value
        cls = Range("F" & i).Value
        chg = cls - opn
        pchg = chg / opn
        
        Range("I" & res_trow).Value = tckr
        Range("L" & res_trow).Value = t_vol
        Range("J" & res_trow).Value = chg
        Range("K" & res_trow).Value = pchg
        
        If Range("J" & res_trow) >= 0 Then
            Range("J" & res_trow).Interior.Color = RGB(0, 250, 0)
        Else
            Range("J" & res_trow).Interior.Color = RGB(250, 0, 0)
        End If
        
        ' Search for greatest total volume change
        
        If t_vol > m_vol Then
            m_vol = t_vol
            Range("P4").Value = tckr
            Range("Q4").Value = m_vol
        End If
        
       
        ' Search for greatest increase and greatest decrease
        
        If pchg > max Then
            max = pchg
            Range("P2").Value = tckr
            Range("Q2").Value = max
        ElseIf pchg < min Then
            min = pchg
            Range("P3").Value = tckr
            Range("Q3").Value = min
        End If
       
        res_trow = res_trow + 1
        t_vol = 0
        tckr_row = 1
    
    Else
        tckr_row = tckr_row + 1
        t_vol = t_vol + Cells(i, 7).Value
    
    End If
  
          

  Next i
  
   
' Find last results row
   
res_lrow = Cells(Rows.Count, 9).End(xlUp).Row

' Format results space

Range("J2", "J" & res_lrow).NumberFormat = "0.00"
Range("K2", "K" & res_lrow).NumberFormat = "0.00%"
Range("Q2", "Q3").NumberFormat = "0.00%"
Range("Q4").NumberFormat = "0,000"

Range("I1:Q4").Columns.AutoFit

Application.ScreenUpdating = True

End Sub
