Option Explicit
Dim Title,Input,i,OutPut
Title = "Oligo Sequence" 
Input = InputBox("GUA",Title,"GUA") 
If input <> "" Then 
    For i = 1 To Len(Input) 
		if i=1 then
			OutPut = "HO-r," & OutPut &  Mid(Input,i,1) & ".p/r," 
		elseif i<Len(Input) then
			OutPut = OutPut &  Mid(Input,i,1) & ".p/r," 
		elseif i=Len(Input) then
			OutPut = OutPut &  Mid(Input,i,1) & "-OH"
		else
		end if
        'OutPut = OutPut &  Mid(Input,i,1) & ".p/r," 
    Next 
End If
MsgBox OutPut,vbInformation,Title
Inputbox "The String " & chr(34)&  Input & chr(34) &" is converted to ",Title,OutPut