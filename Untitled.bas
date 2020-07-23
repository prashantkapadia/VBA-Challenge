Attribute VB_Name = "Module4"
Sub Final()

'' Insterting sources column headers in colums I , J, K,L

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Changed"
Range("K1").Value = "Percentage Changed"
Range("L1").Value = "Total Stock Volume"

For Each ws In Worksheets
ws.Activate

Dim openStock As Double
Dim closeStock As Double
Dim totalStock As Double

totalStock = 0
Dim P_Cent As Long


Dim ticker As String
Dim outputRow As Double
outputRow = 2
Dim lastRow As Double

'Counting  total rows
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ticker = ws.Cells(2, 1).Value
openStock = ws.Cells(2, 3).Value

Dim i As Double
i = 2

'Do While i < lastRow + 1
For i = 2 To lastRow
    If ticker = ws.Cells(i, 1).Value Then
                 
      totalStock = totalStock + Cells(i, 7).Value
    
                   
        Else
                
        ws.Cells(outputRow, 9).Value = ticker
                   
        closeStock = ws.Cells(i - 1, 6).Value
        ws.Cells(outputRow, 10).Value = closeStock - openStock
        
        openStock = ws.Cells(i, 3).Value
        ws.Cells(outputRow, 12).Value = totalStock
         
         'ws.Cells(outputRow, 11).Value = Round((closeStock - openStock) / openStock, 2)
                
         If openStock = 0 Then
         
            openStock = ws.Cells(2, 3).Value
         
            Cells(outputRow, 11).Value = Round((closeStock - openStock) / openStock, 2)
         
            Else
         
                    Cells(outputRow, 11).Value = Round((closeStock - openStock) / openStock, 2)
         
        End If
                
                
                
                
                
          If ws.Cells(outputRow, 11).Value <= 0 Then
          
            Cells(outputRow, 11).Interior.ColorIndex = 3
            
           Else
           
            Cells(outputRow, 11).Interior.ColorIndex = 4
            
           
            End If
        
        
        
        'Reset values
      
        ticker = ws.Cells(i, 1).Value
        totalStock = ws.Cells(i, 7).Value
               
         outputRow = outputRow + 1
        
         
    End If
  '  i = i + 1
    
    
'Loop
Next i

    ws.Cells(outputRow, 9).Value = ticker
    ws.Cells(outputRow, 12).Value = totalStock
    totalStock = ws.Cells(i, 7).Value
    
    
Cells(2, 15).Value = "Greated % Increase"
Cells(3, 15).Value = "Greated % Decrease"
Cells(4, 15).Value = "Greated Toal Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Volume"
    
   
  Dim rng As Range
  Dim Sumrange As Range
  
  Dim dblmax As Double
  Dim dblmin As Double
  Dim Grt_total As Double
  
  Dim T_increase As String
  Dim T_Range As Range
  
  
  
  Dim T_decrease As Long
  Dim Grt_vol As String
  


Set rng = ws.Range("K:K")
Set Sumrange = ws.Range("L:L")

  dblmax = Application.WorksheetFunction.Max(rng)
  dblmin = Application.WorksheetFunction.Min(rng)
  Grt_total = Application.WorksheetFunction.Max(Sumrange)
  
   Cells(2, 17).Value = dblmax
   Cells(3, 17).Value = dblmin
   Cells(4, 17).Value = Grt_total
   
   Dim Mx_Ticker As String
   Dim Mn_Ticker As String
   Dim Great_total As String
   
   
   
   

'Set T_Range = ws.Range("I:K")

'T_increase = Application.WorksheetFunction.VLookup(ws.Range("Q2").Value, T_Range, 8, False)

            For j = 2 To lastRow
            
            If Cells(j, 11).Value = dblmax Then
            
                Mx_Ticker = Cells(j, 9).Value
                
                Cells(2, 16).Value = Mx_Ticker
             
             ElseIf Cells(j, 11).Value = dblmin Then
             
                Mn_Ticker = Cells(j, 9).Value
                
                Cells(3, 16).Value = Mn_Ticker
             
            ElseIf Cells(j, 12).Value = Grt_total Then
             
                Great_total = Cells(j, 9).Value
                
                Cells(4, 16).Value = Great_total
             
            End If
                

     
            Next j
     
    
Next ws


End Sub



