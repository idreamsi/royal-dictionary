Attribute VB_Name = "Ansi2UnicodeModule"
Private Function UniChr(Chr1 As String) As String
Dim Uni As String
Select Case Chr1
    Case Is = "Å"
         Uni = "Ÿæ"
    
    Case Is = "÷"
         Uni = "ÿ∂"
    
    Case Is = "’"
         Uni = "ÿµ"
    
    Case Is = "À"
         Uni = "ÿ´"
    
    Case Is = "ﬁ"
         Uni = "ŸÇ"
    
    Case Is = "›"
         Uni = "ŸÅ"
    
    Case Is = "€"
         Uni = "ÿ∫"
    
    Case Is = "⁄"
         Uni = "ÿπ"
    
    Case Is = "Â"
         Uni = "Ÿá"
    
    Case Is = "Œ"
         Uni = "ÿÆ"
    
    Case Is = "Õ"
         Uni = "ÿ≠"
    
    Case Is = "Ã"
         Uni = "ÿ¨"
    
    Case Is = "ç"
         Uni = "⁄Ü"
    
    Case Is = "‘"
         Uni = "ÿ¥"
    
     Case Is = "”"
         Uni = "ÿ≥"
    
    Case Is = "Ì"
         Uni = "Ÿä"
    
    Case Is = "»"
         Uni = "ÿ®"
    
    Case Is = "·"
         Uni = "ŸÑ"

    
    Case Is = "«"
         Uni = "ÿß"
    
    Case Is = " "
         Uni = "ÿ™"
    
    Case Is = "‰"
         Uni = "ŸÜ"
    
    Case Is = "„"
         Uni = "ŸÖ"
    
    Case Is = "ﬂ"
         Uni = "ŸÉ"
    
    Case Is = "ê"
         Uni = "⁄Ø"
    
    Case Is = "é"
         Uni = "⁄ò"
    
    Case Is = "Ÿ"
         Uni = "ÿ∏"
    
    Case Is = "ÿ"
         Uni = "ÿ∑"
    
    Case Is = "“"
         Uni = "ÿ≤"
    
    Case Is = "—"
         Uni = "ÿ±"
    
    Case Is = "–"
         Uni = "ÿ∞"
    
    Case Is = "œ"
         Uni = "ÿØ"
    
    Case Is = "∆"
         Uni = "ÿ¶"
    
    Case Is = "Ê"
         Uni = "Ÿà"
    
    Case Is = "°"
         Uni = "ÿå"
    
    Case Is = "¬"
         Uni = "ÿ¢"

    
    Case Is = "¡"
         Uni = "ÿ°"
    
    Case Is = "ƒ"
         Uni = "ÿ§"
    Case Is = "Û"
         Uni = "Ÿé"
      
    Case Is = ""
         Uni = "Ÿã"
      
    Case Is = "ı"
         Uni = "Ÿ"
      
    Case Is = "Ò"
         Uni = "Ÿå"
      
    Case Is = "˙"
         Uni = "Ÿí"
      
    Case Is = "ˆ"
         Uni = "Ÿê"
      
    Case Is = "Ú"
         Uni = "Ÿç"
      
    Case Is = "´"
         Uni = "¬´"
      
      Case Is = "ª"
         Uni = "¬ª"
      
      Case Is = "í"
         Uni = "‚Äô"
    
      Case Is = "ë"
         Uni = "‚Äò"
    
      Case Is = "î"
         Uni = "‚Ä"
    
      Case Is = "ì"
         Uni = "‚Äú"

      Case Is = "√"
         Uni = "ÿ£"

       Case Is = "≈"
         Uni = "ÿ•"

      Case Is = "◊"
         Uni = "√ó"
  
      Case Is = "˜"
         Uni = "√∑"
      
      Case Is = "ø"
        Uni = "ÿü"
      
      Case Is = vbNewLine
        Uni = vbNewLine
           
      Case Else
        Uni = Chr1
End Select
UniChr = Uni
End Function
Public Function Ansi2Unicode(ANSI As String) As String
Dim Uni As String
For i = 1 To Len(ANSI)
    l$ = Mid(ANSI, i, 1)
    Uni = UniChr(CStr(l$))
    Ansi2Unicode = Ansi2Unicode & Uni
Next
End Function

