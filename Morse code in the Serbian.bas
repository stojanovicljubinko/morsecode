Attribute VB_Name = "Morse code in the Serbian "
Sub MorseToSerbian_Click()

   Dim MorseConverted As String
   
   Dim x, x1 As Integer
  
    'pronalazimo zadnju praznu celiju u redu 1 - uneti podaci
    lColReq = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'pronalazimo zadnju praznu celiju u redu 4 - morzov kod
    lColMorse = Cells(4, Columns.Count).End(xlToLeft).Column + 1
    
    lColWord = Cells(3, Columns.Count).End(xlToLeft).Columns
    
    
    For x = 1 To lColReq
    
        'prolazak kroz prvi red dok ne naidjemo na prazno polje
        If IsEmpty(Cells(1, x).Text) = True Then
           
         Exit For
      
       End If
    
         'ukoliko je vrednost pronadjena u prvom red, prolazimo kroz 4 red. da bi smo pronasli zamenski karakter - morzov kod u tekst
         For x1 = 1 To lColMorse
               
                If IsEmpty(Cells(4, x1).Text) = True Then
                   Exit For
                End If
               
               'ukoliko je vrednost u 4. redu podudara sa tekstom, koji smo pronasli, spojicemo odgovarajucu vrednost u 3. redu u izlaznu vrednost
               If Cells(4, x1).Text = Cells(1, x).Text Then
                
                 MorseConverted = MorseConverted & Cells(3, x1).Text
               End If
               
               If Cells(3, x1).Text = Cells(1, x).Text Then
               
                MorseConverted = MorseConverted & Cells(4, x1).Text
               End If
        Next x1
        
    Next x
       
    'prikaz rezultata, pretvoren
    MsgBox (MorseConverted)
End Sub
