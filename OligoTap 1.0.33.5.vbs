'inputs for "varriable bases" (vb) 1-3
vb1C=0
vb1O=0
vb1H=1
vb1N=0
vb1P=0
vb1S=0
vb1F=0
'-----
vb2C=0
vb2O=0
vb2H=1
vb2N=0
vb2P=0
vb2S=0
vb2F=0
'-----
vb3C=0
vb3O=0
vb3H=1
vb3N=0
vb3P=0
vb3S=0
vb3F=0
'-----


'the interface code whichsets the user options              
'--------------------------------------------    
'modified nucleosides match the 2021 MODOMICS data base           
Dim objFolder, objItem, objShell, fragfilepath, fragsfilepaths(), filevar(), files(), filenames(),filepaths(), tempfile(), filelines(),  mgf_check(), begin_ion_indices(), end_ion_indices(), oligofragments(), sequences(), oligos(), oligo_1(), oligoname(), tempfile_report, oligoprecursors()          
'Dim oligo()
'begin calculator definitions              
Dim oligofragment_zs(),oligofragment_ints()
dim alpha, beta, o_array0, o_array1, nucleotide, nucleotides, sugar, link, base, var, var1, a(), var2, var3, var4, var5, var6, test               
dim aC(), aH(), aF(), aN(), aO(), aP(), a_S()               
dim abC(), abH(), abF(), abN(), abO(), abP(), ab_S()               
dim message   
dim small_dalton, big_dalton
small_dalton=0.97
big_dalton=1.03            
'end calculator definitions              
' error handling               
On Error Resume Next               
SelectFolder = vbNull               
' Create a dialog object               
Set objShell  = CreateObject( "Shell.Application" )               
Set objFolder = objShell.BrowseForFolder( 0, "Select Folder for the script to generate theoritical fragment tables and a report text file", 0, myStartFolder )               
' Return the path of the selected folder               
If IsObject( objfolder ) Then SelectFolder = objFolder.Self.Path               
fragfilepath = objFolder.Self.Path              
' Standard housekeeping 
objShell.Close        
Set objFolder = Nothing               
Set objshell  = Nothing               
On Error Goto 0              
Set wShell=CreateObject("WScript.Shell")               
Set oExec=wShell.Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")               
oligosfile = oExec.StdOut.ReadLine  
Set oExec=Nothing
Set wShell = Nothing
Set oExec = Nothing             
Dim fsoligo, MyoligoFile               
Set fsoligo = CreateObject("Scripting.FileSystemObject")               
Set MyoligoFile= fsoligo.OpenTextFile(oligosfile, 1)                
MyoligoFile.ReadAll              
ij=MyoligoFile.Line              
MyoligoFile.Close  
Set MyoligoFile = Nothing          
ij_1=ij-1              
'ReDim oligos(ij_1)              
'ReDim oligo_1(ij_1)              
'ReDim oligoname(ij_1)              
'ReDim filevar(ij_1)              
'ReDim fragsfilepaths(ij_1)              
Set MyoligoFile= fsoligo.OpenTextFile(oligosfile, 1)   
'MsgBox "oligos magnitude="&CStr(ij_1)           
c_line_counter=0     
c_oligo_sequence_count=0
line_of_search_sequences=""
Do While c_line_counter<ij
	line_of_search_sequences=MyoligoFile.ReadLine
	multi_line_oligo=InStr(1,line_of_search_sequences,"+",0)
	'multi_line_oligo_line_count=1
	long_oligo=""
	long_oligo_build=""
	if multi_line_oligo>0 then
		long_oligo=Split(line_of_search_sequences,"+")
		Do While multi_line_oligo>0
			'oligos(c_line_counter)=MyoligoFile.ReadLine
			line_of_search_sequences=MyoligoFile.ReadLine
			long_oligo_build=Split(line_of_search_sequences,"+")
			long_oligo(0)=long_oligo(0)&long_oligo_build(0)
			'multi_line_oligo_line_count=multi_line_oligo_line_count+1
			multi_line_oligo=InStr(1,line_of_search_sequences,"+",0)
			c_line_counter=c_line_counter+1
		loop
		'MsgBox	"long_oligos(c) sequence and name=" & long_oligo(0) & "sequence index="&CStr(c_oligo_sequence_count)
		oligo_2=Split(long_oligo(0),"=")
		'MsgBox "long oligo_c(0) sequence="&CStr(oligo_2(0))
		'oligo_1(c_oligo_sequence_count)=oligo_2(0)              
		'oligoname(c_oligo_sequence_count)=oligo_2(1)    
		'fragsfilepaths(c_oligo_sequence_count)=fragfilepath		
		'c_line_counter=c_line_counter+1
		'c_oligo_sequence_count=c_oligo_sequence_count+1
		'MsgBox "loop counter end of loop_long="&CStr(c_line_counter)
	else
		'MsgBox	"short_oligos(c) sequence and name="&CStr(line_of_search_sequences)&"sequence index="&CStr(c_oligo_sequence_count)
		oligo_2=Split(line_of_search_sequences,"=")
		'MsgBox "short=oligo_c(0) sequence="&CStr(oligo_2(0))
		'ReDim Preserve oligo_1(c_oligo_sequence_count)
		'oligo_1(c_oligo_sequence_count)=oligo_2(0)              
		'ReDim Preserve oligoname(c_oligo_sequence_count)
		'oligoname(c_oligo_sequence_count)=oligo_2(1)                          
		'ReDim Preserve fragsfilepaths(c_oligo_sequence_count)
		'fragsfilepaths(c_oligo_sequence_count)=fragfilepath  
		'ReDim Preserve filevar(c_oligo_sequence_count)
		'c_line_counter=c_line_counter+1
		'c_oligo_sequence_count=c_oligo_sequence_count+1
		'MsgBox "input file lines counter magnitude_short="&CStr(c_line_counter)
	end if
	ReDim Preserve oligos(c_oligo_sequence_count)
	oligos(c_oligo_sequence_count)=oligo_2(0) & oligo_2(1)
	ReDim Preserve oligo_1(c_oligo_sequence_count)
	oligo_1(c_oligo_sequence_count)=oligo_2(0)
	ReDim Preserve oligoname(c_oligo_sequence_count)
	oligoname(c_oligo_sequence_count)=oligo_2(1)                          
	ReDim Preserve fragsfilepaths(c_oligo_sequence_count)
	fragsfilepaths(c_oligo_sequence_count)=fragfilepath  
	ReDim Preserve filevar(c_oligo_sequence_count)
	c_oligo_sequence_count=c_oligo_sequence_count+1
	c_line_counter=c_line_counter+1
loop



'MsgBox oligo_1(0)+oligo_1(1)+oligo_1(2)+oligo_1(3)
'MsgBox CStr(TypeName(oligo_1))
MyoligoFile.Close
Set MyoligoFile=Nothing  
Set fsoligo=Nothing          
Dim fso, f3, f1, fc, s, f2, f4, f5, f6             
Set fso = CreateObject("Scripting.FileSystemObject")               
Set objShell  = CreateObject( "Shell.Application" )               
Set f4 = objShell.BrowseForFolder( 0, "Select Folder containing .mgf files", 0, myStartFolder )                           
Set fc = fso.GetFolder(f4.Self.Path).Files         
'possibly input a user input option

      
filelength = 0              
For Each f1 in fc               
    s = s & f1.name                
    s = s &   "<BR>"                
    varriable=Split(f1,"/")              
    mgf_check_length=0              
    for each element in varriable              
        mgf_check_length=mgf_check_length+1              
    Next              
    mgf_check_length_1=0              
    mgf_check_length_1=mgf_check_length-1              
    if varriable(mgf_check_length_1) = "temp/" then              
    else              
        final_mgf_element=Split(varriable(mgf_check_length_1),".")              
        if final_mgf_element(1) = "mgf" then              
            filelength=filelength+1              
        else              
        end if              
    end if                  
Next               
d=filelength-1              
ReDim files(d)              
ReDim filenames(d)              
ReDim filepaths(d)              
ReDim tempfile(d)              
ReDim filelines(d)              
If (fso.FolderExists(fragfilepath&"\temp")) Then               
Else               
    fso.CreateFolder (fragfilepath&"\temp")              
End If               
c7=0              
For Each f1 in fc              
    If f1 = fragfilepath&"\temp" Then               
    Else               
        files(n)=fragfilepath&"\"&f1.name              
        filenames(n)=f1.name              
        Set f2 = fso.GetFile(f1)               
        f2.Copy (fragfilepath&"\temp\"&f1.name)               
        tempfile(c7)=fragfilepath&"\temp\"&f1.name              
    End If                
    c7=c7+1              
    s=c              
Next              
UserInput_precursor = InputBox("Enter acceptable m/z range of precursor isolation(s):"&Chr(13)+"acceptable m/z range"&Chr(13)+"e.g:3","m/z tolerance","3",default)               
ppmrange_precursor=CDbl(UserInput_precursor)              
ppmrange2_precursor=CDbl(ppmrange_precursor/CDbl(2))              
UserInput = InputBox("Enter ppm tolerance of fragment assignments:"&Chr(13)+"acceptable ppm range"&Chr(13)+"e.g:20","m/z tolerance","20",default)               
ppmrange=CDbl(UserInput)              
ppmrange2=ppmrange/2000000              
UserInput_max_charge_per_base = InputBox("Enter max charge of database precursor(s) to search for:"&Chr(13)+"max charge"&Chr(13)+"e.g:5","Charge","5",default)               
max_charge_per_base=CDbl(UserInput_max_charge_per_base) 
UserInput=InputBox("Enter max charge of fragment to search for:","fragment charge states","3",default)
max_charge_fragment=UserInput       


charge_comparison=InputBox("Select charge-isotope strategy"&Chr(13)+"y=monoisotopic (SNAP) w/charge column"&Chr(13)+"x=charge neutral resolved isotope deconvolution"&Chr(13)+"Anyother input leads to no isotope handling","Isotope Strategy","x",default)
'charge_comparison="y"
      
skipcount=0              
c8=0    
Mytime1=Time               
For Each mgf in tempfile              
    'MsgBox "Mgf read event"             
    Set fso = CreateObject("Scripting.FileSystemObject")               
    Set MyFile= fso.OpenTextFile(tempfile(c8), 1)                
    MyFile.ReadAll              
    i=MyFile.Line              
    MyFile.Close   
	Set MyFile=Nothing
    skipcounter=i-1              
    Set MyFile= fso.OpenTextFile(tempfile(c8), 1)               
    begin_ion_count=0              
    end_ion_count=0              
    skipcount=0              
    For j=1 to skipcounter              
        lineinfo=MyFile.ReadLine              
        if lineinfo = "BEGIN IONS" then              
            begin_ion_index=MyFile.Line              
            ReDim Preserve begin_ion_indices(begin_ion_count)              
            begin_ion_indices(begin_ion_count)=MyFile.Line              
            begin_ion_count=begin_ion_count+1              
        elseif  lineinfo = "END IONS" then              
            end_ion_index=MyFile.Line              
            ReDim Preserve end_ion_indices(end_ion_count)              
            end_ion_indices(end_ion_count)=MyFile.Line              
            end_ion_count=end_ion_count+1              
        else              
        End If              
        skipcount=skipcount+1              
    Next              
    MyFile.Close  
	Set MyFile=Nothing
    skipcount=0              
    c6=0                   
    c4=0             
    oligoprecursor=0          
    For each MSMS in begin_ion_indices              
        'MsgBox "msms read evet and c4="&c4             
        oligofragmentcount=0              
        Set MyFile= fso.OpenTextFile(tempfile(c8), 1)              
        k=begin_ion_indices(c6)-1  
        For j=1 to k              
            MyFile.SkipLine              
        Next              
        ion_lines=end_ion_indices(c6)-begin_ion_indices(c6)              
        For j=1 to ion_lines              
            info=MyFile.ReadLine              
            infoarray=Split(info,"=")              
            items=0              
            For each item in infoarray              
                items=items+1              
            Next              
            if infoarray(0)="PEPMASS" then              
                midstring=Mid(info,1,18)              
                oligomassarray=Split(midstring,"=")      				
                oligoprecursor=CDbl(Mid(oligomassarray(1),1,9))             
                'MsgBox oligoprecursor          
                'ReDim Preserve oligoprecursors()           
            else              
            End If              
            If items=1 and j<ion_lines and info<>"END IONS" then  
				'split_info=Split(info," ")
				info_length=len(info)
				'if split_info_length>9 then
				'c13 is meant to be the counter indicating column of the ion lines e.g. mz, intensity, etc
				c13=0
				For i=1 to info_length
					if mid(info,i,1)="0" or mid(info,i,1)="1" or mid(info,i,1)="2" or mid(info,i,1)="3" or mid(info,i,1)="4" or mid(info,i,1)="5" or mid(info,i,1)="6" or mid(info,i,1)="7" or mid(info,i,1)="8" or mid(info,i,1)="9" or mid(info,i,1)="." then
						if c13=0 then
							m_z=m_z&mid(info,i,1)
							'MsgBox "carry="&carry
							'oligofragment=CDbl(midstring)              
							'oligofragmentcount=oligofragmentcount+1              
							'ReDim Preserve oligofragments(oligofragmentcount)              
							'oligofragments(oligofragmentcount)=oligofragment
						elseif c13=1 then
							m_z_int=m_z_int&mid(info,i,1)
						elseif c13=2 then
							m_z_z=m_z_z&mid(info,i,1)
						else
							'oligofragment=carry
							'carry=nothing
							'c13=c13+1
						end If
					'elseif mid(info,i,1)="0" or mid(info,i,1)="1" or mid(info,i,1)="2" or mid(info,i,1)="3" or mid(info,i,1)="4" or mid(info,i,1)="5" or mid(info,i,1)="6" or mid(info,i,1)="7" or mid(info,i,1)="8" or mid(info,i,1)="9" or mid(info,i,1)="." then
					'	carry=carry&mid(split_info(0),i,1)
					else
						c13=c13+1
					end if
				Next
                'midstring=Mid(info,1,9)              
                oligofragment=CDbl(m_z)  
				m_z=""
				oligofragment_int=CDbl(m_z_int)
				m_z_int=""
				if charge_comparison="y" then 
					If IsNumeric(m_z_z) Then
						else
						'...do something
						MsgBox "Common issue is inappropriate peak picking, check spectra are SNAPed with charge assignemnt"&Chr(13)&"all ions must assigned charge in MGF file(s)"
						WScript.Quit()
						WSH.Quit()
					End If
					'MsgBox CStr(m_z_z)
					oligofragment_z=CDbl(m_z_z)
					m_z_z=""
				Else
				end if
                oligofragmentcount=oligofragmentcount+1              
                ReDim Preserve oligofragments(oligofragmentcount)              
                oligofragments(oligofragmentcount)=oligofragment
				ReDim Preserve oligofragments_ints(oligofragmentcount)
				oligofragments_ints(oligofragmentcount)=oligofragment_int
				ReDim Preserve oligofragments_zs(oligofragmentcount)
				oligofragments_zs(oligofragmentcount)=oligofragment_z
                else              
                End If           
        Next              
        'only for the first MSMS of the mgf file set generate text files of the oligo database for subsequent MSMS dont open/close write the files agian                      
        ccc=0          
        For each oligonucleotide in oligos              
            'ccc=0             
            'uniquestructureset=Inputbox("Enter chemical formula for variable base cut at the sugar 1' position. e.g:r,U2.p (MODOMICS Code U2)","Variable Base mods input","C-4,H-3,F-0,N-2,O-1,P-0,S-1",default)              
            'MsgBox "ccc="&ccc             
            filevar(ccc) = fragfilepath+"\"+"oligo_"+oligoname(ccc)+" "+Cstr(ccc)+".txt"              
            'the source code which translates the oligo string to a text file              
            '--------------------------------------------              
            'define arrays for the number of constituent atoms in the w fragments?               
            'difine array for a-b fragment to hold the number of each element in the fragment               
            'oligonucleotide with delimiters to devide the oligo at the building blocks               
            'the 5' and 3' ends are separated by "-"               
            'the ends can be OH , me or ph for oh , methyl or phosphate end               
            'the sugar group is separatred by "," and the r denotes ribose, 2'F ribose as f, 2'o-methyl ribose as rm, 2' o-ethyl-o-methyl ribose as rmoe              
            'the linker is separated by "." and will include phosphorous and sulfur as the core of the phosphate backbone and/or phosphothiorate               
            'the nucleobase is separated by / and the bases will be definied               
            oligo = CStr(oligo_1(ccc))              
            'define the molecule as having zero atoms of each type initially               
            MC=0               
            MH=0               
            MF=0               
            MN=0               
            MO=0               
            MP=0               
            MS=0               
            MC1=0               
            MH1=0               
            MF1=0               
            MN1=0               
            MO1=0               
            MP1=0               
            MS1=0               
            '---------------------------------------------------------------------'               
            'split the user input string describing the oligonucleotide sequence               
            o_array0 = Split(oligo,"-",-1,1)               
            'a test message box to print the sequence with teh ends cut off               
            o_array1 = Split(o_array0(1),"/",-1,1)               
            'count eac of the nucleotides in the sequence               
            i=0               
            For Each nucleotide In o_array1               
                i=i+1               
            Next               
            max_charge_index=CInt(max_charge_per_base)       
            j=i-1              
            'update the fragment array with the new found length of the sequence               
            '----------- alpha ions (used to construct other ions)              
                ReDim negchargestates(i)              
                ReDim poschargestates(i)              
                ReDim aC(i)               
                ReDim aO(i)               
                ReDim aH(i)               
                ReDim aF(i)               
                ReDim aN(i)               
                ReDim aP(i)               
                ReDim a_S(i)               
                ReDim a_mz(i)               
            '----------- [a- base] ions              
                ReDim abC(i)               
                ReDim abH(i)               
                ReDim abF(i)               
                ReDim abN(i)               
                ReDim abO(i)               
                ReDim abP(i)               
                ReDim ab_S(i)               
                ReDim ab_mz(i)               
            '----------- b ions              
                ReDim bC(i)               
                ReDim bH(i)               
                ReDim bF(i)               
                ReDim bN(i)               
                ReDim bO(i)               
                ReDim bP(i)               
                ReDim bS(i)               
                ReDim b_mz(i)               
            '----------- c ions              
                ReDim cC(i)               
                ReDim cH(i)               
                ReDim cF(i)               
                ReDim cN(i)               
                ReDim cO(i)               
                ReDim cP(i)               
                ReDim cS(i)               
                ReDim c_mz(i)               
            '----------- d ions              
                ReDim dC(i)               
                ReDim dH(i)               
                ReDim dF(i)               
                ReDim dN(i)               
                ReDim d_O(i)               
                ReDim dP(i)               
                ReDim dS(i)               
                ReDim d_mz(i)               
            '----------- w ions              
                ReDim wC(i)               
                ReDim wH(i)               
                ReDim wF(i)               
                ReDim wN(i)               
                ReDim wO(i)               
                ReDim wP(i)               
                ReDim wS(i)               
                ReDim w_mz(i)               
            '----------- x ions              
                ReDim xC(i)               
                ReDim xH(i)               
                ReDim xF(i)               
                ReDim xN(i)               
                ReDim xO(i)               
                ReDim xP(i)               
                ReDim xS(i)               
                ReDim x_mz(i)               
            '----------- y ions              
                ReDim yC(i)               
                ReDim yH(i)               
                ReDim yF(i)               
                ReDim yN(i)               
                ReDim yO(i)               
                ReDim yP(i)               
                ReDim yS(i)               
                ReDim y_mz(i)               
            '----------- z ions              
                ReDim zC(i)               
                ReDim zH(i)               
                ReDim zF(i)               
                ReDim zN(i)               
                ReDim zO(i)               
                ReDim zP(i)               
                ReDim zS(i)               
                ReDim z_mz(i)               
                ReDim banana(i)               
                '------------a ions              
                ReDim a1C(i)              
                ReDim a1H(i)              
                ReDim a1F(i)              
                ReDim a1N(i)              
                ReDim a1O(i)              
                ReDim a1P(i)              
                ReDim a1S(i)              
                ReDim a1_mz(i)  
				'-------------'
				ReDim percent(j)
				c11=0
				For Each location in percent
					percent(c11)=0
					c11=c11+1
				Next
				'define boolean array for semi-visual sequence mapping
				ReDim ab_location(j)
				ReDim a_location(j)
				ReDim b_location(j)
				ReDim c_location(j)
				ReDim d_location(j)
				ReDim w_location(j)
				ReDim x_location(j)
				ReDim y_location(j)
				ReDim z_location(j)
				'fill the boolean array with zeros by default
				c14=0
				For Each location in ab_location
					ab_location(c14)=0
					a_location(c14)=0
					b_location(c14)=0
					c_location(c14)=0
					d_location(c14)=0
					w_location(c14)=0
					x_location(c14)=0
					y_location(c14)=0
					z_location(c14)=0
					c14=c14+1
				Next
				c14=0
				'MsgBox c11
            'define an array one less than the length of the sequence               
            '---------------------------------------------------------------------'               
            'add the mass of the 5' end to the structure               
            if o_array0(0) = "HO" then               
                'this is intentionally left blank to correct fot eh sugar construction at the 0 index               
                MH=0               
                MC=0               
                MF=0               
                MN=0               
                MO=0               
                MP=0               
                MS=0               
            elseif o_array0(0) = "ph" then              
                MH=1               
                MO=3               
                MP=1               
                MC=0               
                MF=0               
                MN=0               
                MS=0      
			'dual biotin
            elseif o_array0(0)= "bio2" then              
                MH=65              
                MO=13              
                MP=2              
                MC=36              
                MF=0              
                MN=6              
                MS=2              
            else o_array0(0) = "me"              
                MH=2              
                MC=1              
                MF=0              
                MN=0              
                MO=0              
                MP=0              
                MS=0                  
            End if               
            '---------------------------------------------------------------------'               
            'add the mass of the 3' end to the structure               
            if o_array0(2) = "OH" then               
                MH1=1               
                MC1=0               
                MF1=0               
                MN1=0               
                MO1=1               
                MP1=0               
                MS1=0               
                'MsgBox MH               
            elseif o_array0(2) = "ph" then              
                MH1=2               
                MO1=4               
                MP1=1               
                MC1=0               
                MF1=0               
                MN1=0               
                MS1=0     
			'dual biotin
            elseif o_array0(2)= "bio2" then              
                MH1=66              
                MO1=13              
                MP1=2              
                MC1=36              
                MF1=0              
                MN1=6              
                MS1=2              
            else o_array0(2) = "me"              
                MH1=3              
                MC1=1              
                MF1=0              
                MN1=0              
                MO1=0              
                MP1=0              
                MS1=0                  
            End if               
            '---------------------------------------------------------------------'               
            c=0               
            For Each sugar In o_array1               
                var = Split(o_array1(c),",",-1,1)               
                'add the atoms of the sugar to the fragments               
                if c=0 then               
					'addition of the 5' end to the structure I think this needs to be done before anyother fragments are developed from this ion series               
					aC(c) = aC(c)+MC               
					aH(c) = aH(c)+MH               
					aF(c) = aF(c)+MF               
					aN(c) = aN(c)+MN               
					aO(c) = aO(c)+MO               
					aP(c) = aP(c)+MP               
					a_S(c) = a_S(c)+MS               
					bC(c)=bC(c)+MC               
					bH(c)=bH(c)+MH               
					bO(c)=bO(c)+MO               
					bF(c)=bF(c)+MF               
					bN(c)=bN(c)+MN               
					bP(c)=bP(c)+MP               
					bS(c)=bS(c)+MS               
					if var(0) = "r" then                       
						aC(c)=aC(c)+5               
						aH(c)=aH(c)+3               
						aO(c)=aO(c)+3               
						aF(c)=aF(c)+0               
						aN(c)=aN(c)+0               
						aP(c)=aP(c)+0               
						a_S(c)=a_S(c)+0                 
						bC(c)=bC(c)+5               
						bH(c)=bH(c)+7               
						bO(c)=bO(c)+4               
						bF(c)=bF(c)+0               
						bN(c)=bN(c)+0               
						bP(c)=bP(c)+0               
						bS(c)=bS(c)+0                
						cC(c)=bC(c)               
						cH(c)=bH(c)               
						cO(c)=bO(c)               
						cF(c)=bF(c)               
						cN(c)=bN(c)               
						cP(c)=bP(c)               
						cS(c)=bS(c)       
					ElseIF var(0) = "d" then                       
						aC(c)=aC(c)+5               
						aH(c)=aH(c)+3               
						aO(c)=aO(c)+2               
						aF(c)=aF(c)+0               
						aN(c)=aN(c)+0               
						aP(c)=aP(c)+0               
						a_S(c)=a_S(c)+0               
						bC(c)=bC(c)+5               
						bH(c)=bH(c)+7               
						bO(c)=bO(c)+3               
						bF(c)=bF(c)+0               
						bN(c)=bN(c)+0               
						bP(c)=bP(c)+0               
						bS(c)=bS(c)+0                
						cC(c)=bC(c)               
						cH(c)=bH(c)               
						cO(c)=bO(c)               
						cF(c)=bF(c)               
						cN(c)=bN(c)               
						cP(c)=bP(c)               
						cS(c)=bS(c)              
					ElseIF var(0) = "f" then                       
						aC(c)=aC(c)+5               
						aH(c)=aH(c)+2               
						aO(c)=aO(c)+2               
						aF(c)=aF(c)+1               
						aN(c)=aN(c)+0               
						aP(c)=aP(c)+0               
						a_S(c)=a_S(c)+0               
						bC(c)=bC(c)+5               
						bH(c)=bH(c)+6               
						bO(c)=bO(c)+3               
						bF(c)=bF(c)+1               
						bN(c)=bN(c)+0               
						bP(c)=bP(c)+0               
						bS(c)=bS(c)+0                
						cC(c)=bC(c)               
						cH(c)=bH(c)               
						cO(c)=bO(c)               
						cF(c)=bF(c)               
						cN(c)=bN(c)               
						cP(c)=bP(c)               
						cS(c)=bS(c)                     
					ElseIF var(0) = "rm" then                       
						aC(c)=aC(c)+6               
						aH(c)=aH(c)+5               
						aO(c)=aO(c)+3               
						aF(c)=aF(c)+0               
						aN(c)=aN(c)+0               
						aP(c)=aP(c)+0               
						a_S(c)=a_S(c)+0               
						bC(c)=bC(c)+6               
						bH(c)=bH(c)+9               
						bO(c)=bO(c)+4               
						bF(c)=bF(c)+0               
						bN(c)=bN(c)+0               
						bP(c)=bP(c)+0               
						bS(c)=bS(c)+0               
						cC(c)=bC(c)               
						cH(c)=bH(c)               
						cO(c)=bO(c)               
						cF(c)=bF(c)               
						cN(c)=bN(c)               
						cP(c)=bP(c)               
						cS(c)=bS(c)               
					Else var(0) = "rmoe"                       
						aC(c)=aC(c)+8               
						aH(c)=aH(c)+9               
						aO(c)=aO(c)+4               
						aF(c)=aF(c)+0               
						aN(c)=aN(c)+0               
						aP(c)=aP(c)+0               
						a_S(c)=a_S(c)+0                  
						bC(c)=bC(c)+8               
						bH(c)=bH(c)+13               
						bO(c)=bO(c)+4               
						bF(c)=bF(c)+0               
						bN(c)=bN(c)+0               
						bP(c)=bP(c)+0               
						bS(c)=bS(c)+0               
						cC(c)=bC(c)               
						cH(c)=bH(c)               
						cO(c)=bO(c)               
						cF(c)=bF(c)               
						cN(c)=bN(c)               
						cP(c)=bP(c)               
						cS(c)=bS(c)                   
					End if               
					'for fragments other than the a-base the base of the first nucleotide needs to be added               
					var5 = Split(o_array1(c),".",-1,1)               
					var7 = Split(var5(0),",",-1,1)              
					if var7(1) = "A" then                 
					   bC(c)=bC(c)+5               
					   bH(c)=bH(c)+4               
					   bO(c)=bO(c)+0               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+5               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c)               
					ElseIF var7(1) = "C" then               
					   bC(c)=bC(c)+4               
					   bH(c)=bH(c)+4               
					   bO(c)=bO(c)+1               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+3               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c)               
					ElseIF var7(1) = "5C" then               
					   bC(c)=bC(c)+5               
					   bH(c)=bH(c)+7               
					   bO(c)=bO(c)+1               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+3               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c)               
					ElseIF var7(1) = "G" then               
					   bC(c)=bC(c)+5               
					   bH(c)=bH(c)+4               
					   bO(c)=bO(c)+1               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+5               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c)        
				    ElseIF var7(1) = "22G" then               
					   bC(c)=bC(c)+5+2               
					   bH(c)=bH(c)+4+4               
					   bO(c)=bO(c)+1               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+5               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c)    
					ElseIF var7(1) = "3483G" then   
						'this is the first mod creatd in this way 20220305, the second base mod added after 5c where the delta between the standard is added after the standard  
					   bC(c)=bC(c)+5+11               
					   bH(c)=bH(c)+4+15               
					   bO(c)=bO(c)+1+4               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+5+1               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c) 
					ElseIF var7(1) = "T" then               
					   bC(c)=bC(c)+5               
					   bH(c)=bH(c)+5               
					   bO(c)=bO(c)+2               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+2               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c) 
					Else var7(1) = "U"               
					   bC(c)=bC(c)+4               
					   bH(c)=bH(c)+3               
					   bO(c)=bO(c)+2               
					   bF(c)=bF(c)+0               
					   bN(c)=bN(c)+2               
					   bP(c)=bP(c)+0               
					   bS(c)=bS(c)+0               
					   cC(c)=bC(c)               
					   cH(c)=bH(c)               
					   cO(c)=bO(c)               
					   cF(c)=bF(c)               
					   cN(c)=bN(c)               
					   cP(c)=bP(c)               
					   cS(c)=bS(c)                 
					End If               
					if var5(1) = "p" then               
						'each atomic addition here represents the structure fo the ~typically~ phosphate backbone link between sugars               
						'here an standard phosphate backbone and phosphothiorate are included               
						cC(c)=cC(c)+0               
						cH(c)=cH(c)+0               
						cO(c)=cO(c)+2               
						cF(c)=cF(c)+0               
						cN(c)=cN(c)+0               
						cP(c)=cP(c)+1               
						cS(c)=cS(c)+0               
					Else var5(1) = "s"               
						'this is the phosphothiorate option               
						cC(c)=cC(c)+0               
						cH(c)=cH(c)+0               
						cO(c)=cO(c)+1               
						cF(c)=cF(c)+0               
						cN(c)=cN(c)+0               
						cP(c)=cP(c)+1               
						cS(c)=cS(c)+1               
					End if               
					'beep boop bop at index zero addition of one proton fufils structure, many would not ocnsider it a real fragmetn but i think it makes sense ot fiish it here outside the for loop               
					abC(0) = aC(0)+0               
					'addition of proton to the 3' position to resolve where a link would go for downstream fragments the proton is added at the base posiiton in the base calculation section above               
					abH(0) = aH(0)+1               
					abF(0) = aF(0)+0               
					abN(0) = aN(0)+0               
					abO(0) = aO(0)+0               
					abP(0) = aP(0)+0               
					ab_S(0) = a_S(0)+0               
					'at this point the a-b(0) is complete to construct the other fragments addition of the base will be nessicary               
					'add the hydrogen at the end repeat unit               
					bH(0) = bH(0)+1               
					bF(0) = bF(0)+0               
					bN(0) = bN(0)+0               
					bO(0) = bO(0)+0               
					bP(0) = bP(0)+0               
					bS(0) = bS(0)+0               
					'-------------------------------------------------------------               
                else                   
                var = Split(o_array1(c),",",-1,1)               
                if var(0) = "r" then                       
                    d=c-1               
                    aC(c)=aC(d)+5               
                    aH(c)=aH(d)+3               
                    'additonal hydrogens to saturate the previous sugar can be added here               
                    aH(c)=aH(c)+4               
                    aO(c)=aO(d)+3               
                    aF(c)=aF(d)+0               
                    aN(c)=aN(d)+0               
                    aP(c)=aP(d)+0               
                    a_S(c)=a_S(d)+0                  
                    'additional hydrogens to saturate previous sugar not needed for the b ions               
                    bC(c)=bC(d)+5               
                    bH(c)=bH(d)+7               
                    bO(c)=bO(d)+4               
                    bF(c)=bF(d)+0               
                    bN(c)=bN(d)+0               
                    bP(c)=bP(d)+0               
                    bS(c)=bS(d)+0               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)  
                ElseIF var(0) = "d" then                       
                    d=c-1               
                    aC(c)=aC(d)+5               
                    aH(c)=aH(d)+3               
                    'additonal hydrogens to saturate the previous sugar can be added here               
                    aH(c)=aH(c)+4               
                    aO(c)=aO(d)+2               
                    aF(c)=aF(d)+0               
                    aN(c)=aN(d)+0               
                    aP(c)=aP(d)+0               
                    a_S(c)=a_S(d)+0                      
                    bC(c)=bC(d)+5               
                    bH(c)=bH(d)+7               
                    bO(c)=bO(d)+3               
                    bF(c)=bF(d)+0               
                    bN(c)=bN(d)+0               
                    bP(c)=bP(d)+0               
                    bS(c)=bS(d)+0               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)                           
                ElseIF var(0) = "f" then                       
                    d=c-1               
                    aC(c)=aC(d)+5               
                    aH(c)=aH(d)+2               
                    'additonal hydrogens to saturate the previous sugar can be added here               
                    aH(c)=aH(c)+4               
                    aO(c)=aO(d)+2               
                    aF(c)=aF(d)+1               
                    aN(c)=aN(d)+0               
                    aP(c)=aP(d)+0               
                    a_S(c)=a_S(d)+0                      
                    bC(c)=bC(d)+5               
                    bH(c)=bH(d)+6               
                    bO(c)=bO(d)+3               
                    bF(c)=bF(d)+1               
                    bN(c)=bN(d)+0               
                    bP(c)=bP(d)+0               
                    bS(c)=bS(d)+0               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)                      
                ElseIF var(0) = "rm" then                       
                    d=c-1               
                    aC(c)=aC(d)+6               
                    aH(c)=aH(d)+5               
                    'additonal hydrogens to saturate the previous sugar can be added here               
                    aH(c)=aH(c)+4               
                    aO(c)=aO(d)+3               
                    aF(c)=aF(d)+0               
                    aN(c)=aN(d)+0               
                    aP(c)=aP(d)+0               
                    a_S(c)=a_S(d)+0                      
                    bC(c)=bC(d)+6               
                    bH(c)=bH(d)+9               
                    bO(c)=bO(d)+4               
                    bF(c)=bF(d)+0               
                    bN(c)=bN(d)+0               
                    bP(c)=bP(d)+0               
                    bS(c)=bS(d)+0               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)               
                Else var(0) = "rmoe"                       
                    d=c-1               
                    aC(c)=aC(d)+8               
                    aH(c)=aH(d)+9               
                    'additonal hydrogens to saturate the previous sugar can be added here               
                    aH(c)=aH(c)+4               
                    aO(c)=aO(d)+4               
                    aF(c)=aF(d)+0               
                    aN(c)=aN(d)+0               
                    aP(c)=aP(d)+0               
                    a_S(c)=a_S(d)+0                    
                    bC(c)=bC(d)+8               
                    bH(c)=bH(d)+13               
                    bO(c)=bO(d)+4               
                    bF(c)=bF(d)+0               
                    bN(c)=bN(d)+0               
                    bP(c)=bP(d)+0               
                    bS(c)=bS(d)+0               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)                     
                End if               
                d=c-1               
                var5 = Split(o_array1(d),".",-1,1)               
                'final isolation of the oligonucleotide element from the string by isolation with the other delimiter on the other side of the element                
                'adding the previous linker d ot the current fragment (this section of the code is in the index>=1 so it should be correct)[remembering that indices start at zero]               
                if var5(1) = "p" then               
                    'each atomic addition here represents the structure fo the ~typically~ phosphate backbone link between sugars               
                    'here an standard phosphate backbone and phosphothiorate are included               
                    aC(c)=aC(c)+0               
                    aH(c)=aH(c)+1               
                    aO(c)=aO(c)+3               
                    aF(c)=aF(c)+0               
                    aN(c)=aN(c)+0               
                    aP(c)=aP(c)+1               
                    a_S(c)=a_S(c)+0               
                    bC(c)=aC(c)               
                    bH(c)=aH(c)               
                    bO(c)=aO(c)               
                    bF(c)=aF(c)               
                    bN(c)=aN(c)               
                    bP(c)=aP(c)               
                    bS(c)=a_S(c)               
                    'account for the previous linker               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)                 
                Else var5(1) = "s"               
                    'this is the phosphothiorate option               
                    aC(c)=aC(c)+0               
                    aH(c)=aH(c)+1               
                    aO(c)=aO(c)+2               
                    aF(c)=aF(c)+0               
                    aN(c)=aN(c)+0               
                    aP(c)=aP(c)+1               
                    a_S(c)=a_S(c)+1               
                    bC(c)=aC(c)               
                    bH(c)=aH(c)               
                    bO(c)=aO(c)               
                    bF(c)=aF(c)               
                    bN(c)=aN(c)               
                    bP(c)=aP(c)               
                    bS(c)=a_S(c)               
                    'account for the previous linker               
                    cC(c)=bC(c)               
                    cH(c)=bH(c)               
                    cO(c)=bO(c)               
                    cF(c)=bF(c)               
                    cN(c)=bN(c)               
                    cP(c)=bP(c)               
                    cS(c)=bS(c)                 
                    'MsgBox "s tree"               
                End if               
                'taking the previous base and add it to the a fragment               
                var7 = Split(var5(0),",",-1,1)              
                if var7(1) = "A" then                 
                   aC(c)=aC(c)+5               
                   aH(c)=aH(c)+4               
                   aO(c)=aO(c)+0               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+5               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0               
                ElseIF var7(1) = "C" then               
                   aC(c)=aC(c)+4               
                   aH(c)=aH(c)+4               
                   aO(c)=aO(c)+1               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+3               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0               
                ElseIF var7(1) = "5C" then               
                   aC(c)=aC(c)+5               
                   aH(c)=aH(c)+7               
                   aO(c)=aO(c)+1               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+3               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0               
                ElseIF var7(1) = "G" then               
                   aC(c)=aC(c)+5               
                   aH(c)=aH(c)+4               
                   aO(c)=aO(c)+1               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+5               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0       
                ElseIF var7(1) = "22G" then               
                   aC(c)=aC(c)+5+2               
                   aH(c)=aH(c)+4+4               
                   aO(c)=aO(c)+1               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+5               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0                      
                ElseIF var7(1) = "3483G" then               
                   aC(c)=aC(c)+5+11               
                   aH(c)=aH(c)+4+15               
                   aO(c)=aO(c)+1+4               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+5+1               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0  
				ElseIF var7(1) = "T" then               
                   aC(c)=aC(c)+5               
                   aH(c)=aH(c)+5               
                   aO(c)=aO(c)+2               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+2               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0    
                Else var7(1) = "U"               
                   aC(c)=aC(c)+4               
                   aH(c)=aH(c)+3               
                   aO(c)=aO(c)+2               
                   aF(c)=aF(c)+0               
                   aN(c)=aN(c)+2               
                   aP(c)=aP(c)+0               
                   a_S(c)=a_S(c)+0                 
                End If               
                'take the current base and add to the current b fragment               
                var5 = Split(o_array1(c),".",-1,1)               
                var7 = Split(var5(0),",",-1,1)              
                if var7(1) = "A" then                 
                   bC(c)=aC(c)+5               
                   bH(c)=aH(c)+4               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+0               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+5               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)               
                ElseIF var7(1) = "C" then               
                   bC(c)=aC(c)+4               
                   bH(c)=aH(c)+4               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+1               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+3               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)               
                ElseIF var7(1) = "5C" then               
                   bC(c)=aC(c)+5               
                   bH(c)=aH(c)+7               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+1               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+3               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)              
                ElseIF var7(1) = "G" then               
                   bC(c)=aC(c)+5               
                   bH(c)=aH(c)+4               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+1               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+5               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)     
               ElseIF var7(1) = "22G" then               
                   bC(c)=aC(c)+5+2               
                   bH(c)=aH(c)+4+4               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+1               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+5               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)     
                ElseIF var7(1) = "3483G" then               
                   bC(c)=aC(c)+5+11               
                   bH(c)=aH(c)+4+15               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+1+4               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+5+1               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)    
			    ElseIF var7(1) = "T" then               
                   bC(c)=aC(c)+5               
                   bH(c)=aH(c)+5               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+2               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+2               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)    
                Else var7(1) = "U"               
                   bC(c)=aC(c)+4               
                   bH(c)=aH(c)+3               
                   'additonal hydrogens to saturate the previous sugar possibly belongs here               
                   bH(c)=bH(c)+4               
                   bO(c)=aO(c)+2               
                   bF(c)=aF(c)+0               
                   bN(c)=aN(c)+2               
                   bP(c)=aP(c)+0               
                   bS(c)=a_S(c)+0               
                   'the base addition is the same for the b and c ions               
                   cC(c)=bC(c)               
                   cH(c)=bH(c)               
                   cO(c)=bO(c)               
                   cF(c)=bF(c)               
                   cN(c)=bN(c)               
                   cP(c)=bP(c)               
                   cS(c)=bS(c)                 
                End If               
                'since the base addition is the last event in the construction of the b ion it would seem the final addition to complete the c ion must be done after the b ion               
                'add the current linker               
                var5 = Split(o_array1(c),".",-1,1)              
                    if c<j then              
                        if var5(1) = "p" then                
                           'each atomic addition here represents the structure fo the ~typically~ phosphate backbone link between sugars                
                           'here an standard phosphate backbone and phosphothiorate are included                
                           cC(c)=cC(c)+0                
                           cH(c)=cH(c)+0                
                           cO(c)=cO(c)+3                
                           cF(c)=cF(c)+0                
                           cN(c)=cN(c)+0                
                           cP(c)=cP(c)+1                
                           cS(c)=cS(c)+0                
                        Else var5(1) = "s"                
                           'this is the phosphothiorate option                
                           cC(c)=cC(c)+0                
                           cH(c)=cH(c)+0                
                           cO(c)=cO(c)+2                
                           cF(c)=cF(c)+0                
                           cN(c)=cN(c)+0                
                           cP(c)=cP(c)+1                
                           cS(c)=cS(c)+1                
                           'MsgBox "s tree"                
                        End if               
                        else              
                    End If               
                'addition of two protons to cap the two carbons where additonal bonds will be made in the subsequent fragment                   
                abC(c) = aC(c)+0               
                abH(c) = aH(c)+2               
                abF(c) = aF(c)+0               
                abN(c) = aN(c)+0               
                abO(c) = aO(c)+0               
                abP(c) = aP(c)+0               
                ab_S(c) = a_S(c)+0               
                    if c<j then              
                        'add to the 3' position the O and H which matched my structure design i think              
                        bC(c) = bC(c)+0               
                        bH(c) = bH(c)+1               
                        bF(c) = bF(c)+0               
                        bN(c) = bN(c)+0               
                        bO(c) = bO(c)+1               
                        bP(c) = bP(c)+0               
                        bS(c) = bS(c)+0               
                    Else c=j              
                        MC=bC(c)+MC1               
                        MH=bH(c)+MH1               
                        MF=bF(c)+MF1               
                        MN=bN(c)+MN1               
                        MO=bO(c)+MO1              
                        MP=bP(c)+MP1               
                        MS=bS(c)+MS1               
                        Mstr="M - H,    C"+CStr(MC)+"H"+CStr(MH)+"F"+CStr(MF)+"N"+CStr(MN)+"O"+CStr(MO)+"P"+CStr(MP)+"S"+CStr(MS)+" "              
                    End If              
                End If               
                dC(c) = cC(c)+0               
                dH(c) = cH(c)+2               
                dF(c) = cF(c)+0               
                dN(c) = cN(c)+0               
                d_O(c) = cO(c)+1               
                dP(c) = cP(c)+0               
                dS(c) = cS(c)+0               
                a1C(c) = bC(c)-0               
                a1H(c) = bH(c)-0               
                a1F(c) = bF(c)-0               
                a1N(c) = bN(c)-0               
                a1O(c) = bO(c)-1               
                a1P(c) = bP(c)-0               
                a1S(c) = bS(c)-0               
                c=c+1               
            Next               
            j=c               
            c=0               
            d=j-1              
            For Each sugar In o_array1                
                wC(c) = MC-a1C(d)               
                wH(c) = MH-a1H(d)+1               
                wF(c) = MF-a1F(d)               
                wN(c) = MN-a1N(d)               
                wO(c) = MO-a1O(d)              
                wP(c) = MP-a1P(d)               
                wS(c) = MS-a1S(d)               
                xC(c) = wC(c)               
                xH(c) = wH(c)-2              
                xF(c) = wF(c)               
                xN(c) = wN(c)               
                xO(c) = wO(c)-1              
                xP(c) = wP(c)               
                xS(c) = wS(c)               
                yC(c) = MC-cC(d)               
                yH(c) = MH-cH(d)-1               
                yF(c) = MF-cF(d)               
                yN(c) = MN-cN(d)               
                yO(c) = MO-cO(d)               
                yP(c) = MP-cP(d)               
                yS(c) = MS-cS(d)               
                zC(c) = yC(c)               
                zH(c) = yH(c)-1               
                zF(c) = yF(c)               
                zN(c) = yN(c)               
                zO(c) = yO(c)-1               
                zP(c) = yP(c)               
                zS(c) = yS(c)                 
                d=d-1               
                c=c+1               
            Next               
            d=c-1               
            wC(d) = wC(d)               
            wH(d) = wH(d)               
            wF(d) = wF(d)               
            wN(d) = wN(d)               
            wO(d) = wO(d)              
            wP(d) = wP(d)               
            wS(d) = wS(d)                
            '------------------------------------------------------'               
            c=0               
            C_ = 12               
            H = 1.007825               
            F = 18.998403               
            N = 14.003074               
            O = 15.994915               
            P = 30.973763               
            S = 31.972072               
            Mmstr=Cstr(MC*C_+MH*H+MF*F+MN*N+MO*O+MP*P+MS*S)              
            q=1              
            c=0              
            For Each z in negchargestates              
                negchargestates(c)=Cstr((MC*C_+(MH+1-q)*H+MF*F+MN*N+MO*O+MP*P+MS*S)/q)              
                poschargestates(c)=Cstr((MC*C_+(MH+1+q)*H+MF*F+MN*N+MO*O+MP*P+MS*S)/q)              
                q=q+1              
                c=c+1              
            Next              
            q=0              
            c=0              
            For Each sugar In o_array1                
                a1_mz(c)    =CStr(a1C(c)*C_+a1H(c)*H+a1F(c)*F+a1N(c)*N+a1O(c)*O+a1P(c)*P+a1S(c)*S)               
                ab_mz(c)    =CStr(abC(c)*C_+abH(c)*H+abF(c)*F+abN(c)*N+abO(c)*O+abP(c)*P+ab_S(c)*S)               
                b_mz(c)     =CStr(bC(c)*C_+bH(c)*H+bF(c)*F+bN(c)*N+bO(c)*O+bP(c)*P+bS(c)*S)               
                c_mz(c)     =CStr(cC(c)*C_+cH(c)*H+cF(c)*F+cN(c)*N+cO(c)*O+cP(c)*P+cS(c)*S)               
                d_mz(c)     =CStr(dC(c)*C_+dH(c)*H+dF(c)*F+dN(c)*N+d_O(c)*O+dP(c)*P+dS(c)*S)               
                w_mz(c)     =CStr(wC(c)*C_+wH(c)*H+wF(c)*F+wN(c)*N+wO(c)*O+wP(c)*P+wS(c)*S)                
                x_mz(c)     =CStr(xC(c)*C_+xH(c)*H+xF(c)*F+xN(c)*N+xO(c)*O+xP(c)*P+xS(c)*S)               
                y_mz(c)     =CStr(yC(c)*C_+yH(c)*H+yF(c)*F+yN(c)*N+yO(c)*O+yP(c)*P+yS(c)*S)               
                z_mz(c)     =CStr(zC(c)*C_+zH(c)*H+zF(c)*F+zN(c)*N+zO(c)*O+zP(c)*P+zS(c)*S)               
                c=c+1               
            Next              
            if c4=0 then              
                Dim f                
                Set fso = CreateObject("Scripting.FileSystemObject")                
                Set f = fso.OpenTextFile(filevar(ccc), 2, True)                
                c=0               
                q=c+1               
                f.Write oligo&Chr(13)              
                f.Write Mmstr&Chr(13)              
                f.Write Mstr&Chr(13)              
                f.Write "molecular charge states"&Chr(13)               
                c=0              
                g=c+1              
                For Each z in negchargestates              
                    f.Write  negchargestates(c)&",  charge=-"&g&Chr(13)              
                    f.Write  poschargestates(c)&",  charge=+"&g&Chr(13)              
                    c=c+1              
                    g=g+1              
                Next              
                c=0              
                f.Write "all charge states are (1-)"&Chr(13)               
                For Each sugar In o_array1                
                    If q=1 then               
                    ElseIf q<i then               
                        f.Write "a-b"+CStr(q)+Chr(44)+Chr(9)+ab_mz(c)+Chr(44)+Chr(9)+"C"+CStr(abC(c))+"H"+CStr(abH(c))+"F"+CStr(abF(c))+"N"+CStr(abN(c))+"O"+CStr(abO(c))+"P"+CStr(abP(c))+"S"+CStr(ab_S(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                    q=q+1               
                Next               
                c=0               
                q=c+1               
                For Each sugar In o_array1               
                    if q<i then               
                        f.Write "a"+CStr(q)+Chr(44)+Chr(9)+a1_mz(c)+Chr(44)+Chr(9)+"C"+CStr(a1C(c))+"H"+CStr(a1H(c))+"F"+CStr(a1F(c))+"N"+CStr(a1N(c))+"O"+CStr(a1O(c))+"P"+CStr(a1P(c))+"S"+CStr(a1S(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                    q=q+1               
                Next               
                c=0               
                q=c+1               
                For Each sugar In o_array1               
                    if q<i then               
                        f.Write "b"+CStr(q)+Chr(44)+Chr(9)+b_mz(c)+Chr(44)+Chr(9)+"C"+CStr(bC(c))+"H"+CStr(bH(c))+"F"+CStr(bF(c))+"N"+CStr(bN(c))+"O"+CStr(bO(c))+"P"+CStr(bP(c))+"S"+CStr(bS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                    q=q+1               
                Next               
                c=0               
                q=c+1               
                For Each sugar In o_array1               
                    if q<i then                
                        f.Write "c"+CStr(q)+Chr(44)+Chr(9)+c_mz(c)+Chr(44)+Chr(9)+"C"+CStr(cC(c))+"H"+CStr(cH(c))+"F"+CStr(cF(c))+"N"+CStr(cN(c))+"O"+CStr(cO(c))+"P"+CStr(cP(c))+"S"+CStr(cS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                    q=q+1               
                Next               
                c=0               
                q=c+1               
                For Each sugar In o_array1               
                    if q<i then                
                        f.Write "d"+CStr(q)+Chr(44)+Chr(9)+d_mz(c)+Chr(44)+Chr(9)+"C"+CStr(dC(c))+"H"+CStr(dH(c))+"F"+CStr(dF(c))+"N"+CStr(dN(c))+"O"+CStr(d_O(c))+"P"+CStr(dP(c))+"S"+CStr(dS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                    q=q+1               
                Next               
                c=1               
                For Each sugar In o_array1                
                    if c<i then               
                        f.Write "w"+CStr(c)+Chr(44)+Chr(9)+w_mz(c)+Chr(44)+Chr(9)+"C"+CStr(wC(c))+"H"+CStr(wH(c))+"F"+CStr(wF(c))+"N"+CStr(wN(c))+"O"+CStr(wO(c))+"P"+CStr(wP(c))+"S"+CStr(wS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                Next               
                c=1               
                For Each sugar In o_array1                
                    if c<i then               
                        f.Write "x"+CStr(c)+Chr(44)+Chr(9)+x_mz(c)+Chr(44)+Chr(9)+"C"+CStr(xC(c))+"H"+CStr(xH(c))+"F"+CStr(xF(c))+"N"+CStr(xN(c))+"O"+CStr(xO(c))+"P"+CStr(xP(c))+"S"+CStr(xS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                Next               
                c=1               
                For Each sugar In o_array1                
                    if c<i then               
                        f.Write "y"+CStr(c)+Chr(44)+Chr(9)+y_mz(c)+Chr(44)+Chr(9)+"C"+CStr(yC(c))+"H"+CStr(yH(c))+"F"+CStr(yF(c))+"N"+CStr(yN(c))+"O"+CStr(yO(c))+"P"+CStr(yP(c))+"S"+CStr(yS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                Next               
                c=1               
                For Each sugar In o_array1                
                    if c<i then               
                        f.Write "z"+CStr(c)+Chr(44)+Chr(9)+z_mz(c)+Chr(44)+Chr(9)+"C"+CStr(zC(c))+"H"+CStr(zH(c))+"F"+CStr(zF(c))+"N"+CStr(zN(c))+"O"+CStr(zO(c))+"P"+CStr(zP(c))+"S"+CStr(zS(c))+" "&Chr(13)               
                    End If               
                    c=c+1               
                Next               
                f.Close              
            else              
            End If              
            ccc=ccc+1              
            '---              
            'insert peak comparison here          
            Set fso = CreateObject("Scripting.FileSystemObject")        
            Set f6 = fso.OpenTextFile(fragfilepath+"\"+"report.txt", 8, True)                
            '---              
            c5=0          
            'MsgBox max_charge_index      
            For precursor=0 to max_charge_index          
                oligoprecursor_H=CDbl(negchargestates(c5))+CDbl(ppmrange2_precursor)              
                oligoprecursor_L=CDbl(negchargestates(c5))-CDbl(ppmrange2_precursor)         
                if    CDbl(oligoprecursor)<CDbl(oligoprecursor_H) and CDbl(oligoprecursor)>CDbl(oligoprecursor_L) then          
                    f6.Write CStr(tempfile(c8))&Chr(13)        
                    f6.Write "neg"+Chr(44)+Chr(9)+"exp="+CStr(oligoprecursor)+Chr(44)+Chr(9)+"theo="+CStr(negchargestates(c5))+Chr(44)+Chr(9)+oligoname(ccc-1)+Chr(44)+Chr(9)+"charge="+CStr(precursor+1)&Chr(13)        
                    c9=0      
                    d1=0 
					c12=0
					For Each location in percent
						c12=c12+1
					Next
                    For Each sugar In o_array1              
                        d1=c9+1              
                        'a1_mz_n(c)=CDbl(a1_mz(c))     
						'ok i think if i reconstruct the high low ppm range value in the onditional line at 1418 I can have the multiple charge state looped in the comparison
                        'a1_mz_L     =CDbl(a1_mz(c9))-CDbl(a1_mz(c9))*ppmrange2              
                        'a1_mz_H     =CDbl(a1_mz(c9))+CDbl(a1_mz(c9))*ppmrange2              
                        'ab_mz_L     =CDbl(ab_mz(c9))-CDbl(ab_mz(c9))*ppmrange2              
                        'ab_mz_H     =CDbl(ab_mz(c9))+CDbl(ab_mz(c9))*ppmrange2              
                        'b_mz_L     =CDbl(b_mz(c9))-CDbl(b_mz(c9))*ppmrange2              
                        'b_mz_H     =CDbl(b_mz(c9))+CDbl(b_mz(c9))*ppmrange2              
                        'c_mz_L     =CDbl(c_mz(c9))-CDbl(c_mz(c9))*ppmrange2              
                        'c_mz_H     =CDbl(c_mz(c9))+CDbl(c_mz(c9))*ppmrange2              
                        'd_mz_L     =CDbl(d_mz(c9))-CDbl(d_mz(c9))*ppmrange2              
                        'd_mz_H     =CDbl(d_mz(c9))+CDbl(d_mz(c9))*ppmrange2              
                        'w_mz_L     =CDbl(w_mz(c9))-CDbl(w_mz(c9))*ppmrange2              
                        'w_mz_H     =CDbl(w_mz(c9))+CDbl(w_mz(c9))*ppmrange2              
                        'x_mz_L     =CDbl(x_mz(c9))-CDbl(x_mz(c9))*ppmrange2              
                        'x_mz_H     =CDbl(x_mz(c9))+CDbl(x_mz(c9))*ppmrange2              
                        'y_mz_L     =CDbl(y_mz(c9))-CDbl(y_mz(c9))*ppmrange2              
                        'y_mz_H     =CDbl(y_mz(c9))+CDbl(y_mz(c9))*ppmrange2              
                        'z_mz_L     =CDbl(z_mz(c9))-CDbl(z_mz(c9))*ppmrange2              
                        'z_mz_H     =CDbl(z_mz(c9))+CDbl(z_mz(c9))*ppmrange2              
                        'MsgBox a1_mz_n_L           
                        'this is probably the place to put annother for loop to itterate through the fragmtn ion signals      
                        if charge_comparison="y" then
						'c10 is the counter to step through all the read fragment signal lines from the mgf file
							c10=0      
							For Each signal In oligofragments 
								For z=1 to max_charge_fragment
									if oligofragments(c10)<CDbl((a1_mz(c9)-(z-1)*H)/z)+CDbl((a1_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((a1_mz(c9)-(z-1)*H)/z)-CDbl((a1_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                      
										'annotate here       
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "a"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((a1_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(a1C(c9))+"H"+CStr(a1H(c9)-(z-1))+"F"+CStr(a1F(c9))+"N"+CStr(a1N(c9))+"O"+CStr(a1O(c9))+"P"+CStr(a1P(c9))+"S"+CStr(a1S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((a1_mz(c9)-(z-1)*H)/z))/((a1_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                        
										percent(d1)=1
										a_location(d1)=1
									else              
									End if              
									if oligofragments(c10)<CDbl((ab_mz(c9)-(z-1)*H)/z)+CDbl((ab_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((ab_mz(c9)-(z-1)*H)/z)-CDbl((ab_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                  
										'Cmpd(2).Annotations.AddAnnotation ab_mz_L, prime_int*1.2, ab_mz_H, prime_int*1.2, prime_int*1.6, "a-b"+CStr(d), True              
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "ab"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((ab_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(abC(c9))+"H"+CStr(abH(c9)-(z-1))+"F"+CStr(abF(c9))+"N"+CStr(abN(c9))+"O"+CStr(abO(c9))+"P"+CStr(abP(c9))+"S"+CStr(ab_S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((ab_mz(c9)-(z-1)*H)/z))/((ab_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)              
										percent(d1)=1
										ab_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((b_mz(c9)-(z-1)*H)/z)+CDbl((b_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((b_mz(c9)-(z-1)*H)/z)-CDbl((b_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                   
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "b"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((b_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(bC(c9))+"H"+CStr(bH(c9)-(z-1))+"F"+CStr(bF(c9))+"N"+CStr(bN(c9))+"O"+CStr(bO(c9))+"P"+CStr(bP(c9))+"S"+CStr(bS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((b_mz(c9)-(z-1)*H)/z))/((b_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                      
										percent(d1)=1
										b_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((c_mz(c9)-(z-1)*H)/z)+CDbl((c_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((c_mz(c9)-(z-1)*H)/z)-CDbl((c_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                     
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)             
										f6.Write "c"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((c_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(cC(c9))+"H"+CStr(cH(c9)-(z-1))+"F"+CStr(cF(c9))+"N"+CStr(cN(c9))+"O"+CStr(cO(c9))+"P"+CStr(cP(c9))+"S"+CStr(cS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((c_mz(c9)-(z-1)*H)/z))/(c_mz(c9))*1000000),4))&Chr(13)              
										percent(d1)=1
										c_location(d1)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((d_mz(c9)-(z-1)*H)/z)+CDbl((d_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((d_mz(c9)-(z-1)*H)/z)-CDbl((d_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                   
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)       
										f6.Write "d"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((d_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(dC(c9))+"H"+CStr(dH(c9)-(z-1))+"F"+CStr(dF(c9))+"N"+CStr(dN(c9))+"O"+CStr(d_O(c9))+"P"+CStr(dP(c9))+"S"+CStr(dS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((d_mz(c9)-(z-1)*H)/z))/((d_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                     
										percent(d1)=1
										d_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((w_mz(c9)-(z-1)*H)/z)+CDbl((w_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((w_mz(c9)-(z-1)*H)/z)-CDbl((w_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then                 
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "w"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((w_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(wC(c9))+"H"+CStr(wH(c9)-(z-1))+"F"+CStr(wF(c9))+"N"+CStr(wN(c9))+"O"+CStr(wO(c9))+"P"+CStr(wP(c9))+"S"+CStr(wS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((w_mz(c9)-(z-1)*H)/z))/((w_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)              
										'MsgBox "c12="&c12
										'MsgBox "c9="&c9
										percent(c12-c9)=1
										w_location(c12-c9)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((x_mz(c9)-(z-1)*H)/z)+CDbl((x_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((x_mz(c9)-(z-1)*H)/z)-CDbl((x_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then                    
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "x"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((x_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(xC(c9))+"H"+CStr(xH(c9)-(z-1))+"F"+CStr(xF(c9))+"N"+CStr(xN(c9))+"O"+CStr(xO(c9))+"P"+CStr(xP(c9))+"S"+CStr(xS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((x_mz(c9)-(z-1)*H)/z))/((x_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                
										percent(c12-c9)=1
										x_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((y_mz(c9)-(z-1)*H)/z)+CDbl((y_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((y_mz(c9)-(z-1)*H)/z)-CDbl((y_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "y"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((y_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(yC(c9))+"H"+CStr(yH(c9)-(z-1))+"F"+CStr(yF(c9))+"N"+CStr(yN(c9))+"O"+CStr(yO(c9))+"P"+CStr(yP(c9))+"S"+CStr(yS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((y_mz(c9)-(z-1)*H)/z))/((y_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										y_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((z_mz(c9)-(z-1)*H)/z)+CDbl((z_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((z_mz(c9)-(z-1)*H)/z)-CDbl((z_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then            
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "z"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((z_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(zC(c9))+"H"+CStr(zH(c9)-(z-1))+"F"+CStr(zF(c9))+"N"+CStr(zN(c9))+"O"+CStr(zO(c9))+"P"+CStr(zP(c9))+"S"+CStr(zS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((z_mz(c9)-(z-1)*H)/z))/((z_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										z_location(c12-c9)=1
									else              
									end if    
								Next
								'c=c+1      
								c10=c10+1      
							Next      
						Elseif charge_comparison="x" then
							c10=0      
							For Each signal In oligofragments 
								'For z=1 to max_charge_fragment
								if oligofragments(c10)<CDbl((a1_mz(c9)+(1*H)))+CDbl((a1_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((a1_mz(c9)+(1*H)))-CDbl((a1_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                      
									'the second if statmetn below is to check if the signal in the oligofragments has a signal which is between 0.97 and 1.03 daltons lower mass than the signal itself, if false this peak is not a descernable monooisotopic peak in this workflow
									'else if the peak is a descernable monoisotopic peak it may be assigned
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "a"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((a1_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(a1C(c9))+"H"+CStr(a1H(c9)+1)+"F"+CStr(a1F(c9))+"N"+CStr(a1N(c9))+"O"+CStr(a1O(c9))+"P"+CStr(a1P(c9))+"S"+CStr(a1S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(a1_mz(c9)+(1*H)))/(a1_mz(c9)+(1*H))*1000000),4))&Chr(13)                        
										percent(d1)=1
										a_location(d1)=1
									Else
									End if
								else              
								End if              
								if oligofragments(c10)<CDbl((ab_mz(c9)+(1*H)))+CDbl((ab_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((ab_mz(c9)+(1*H)))-CDbl((ab_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                  
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "ab"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((ab_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(abC(c9))+"H"+CStr(abH(c9)+1)+"F"+CStr(abF(c9))+"N"+CStr(abN(c9))+"O"+CStr(abO(c9))+"P"+CStr(abP(c9))+"S"+CStr(ab_S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(ab_mz(c9)+(1*H)))/(ab_mz(c9)+(1*H))*1000000),4))&Chr(13)              
										percent(d1)=1
										ab_location(d1)=1
									Else
									End if
								else              
								end if              
								if oligofragments(c10)<CDbl((b_mz(c9)+(1*H)))+CDbl((b_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((b_mz(c9)+(1*H)))-CDbl((b_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                   
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next	
									if light_logic=true then
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "b"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((b_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(bC(c9))+"H"+CStr(bH(c9)+1)+"F"+CStr(bF(c9))+"N"+CStr(bN(c9))+"O"+CStr(bO(c9))+"P"+CStr(bP(c9))+"S"+CStr(bS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(b_mz(c9)+(1*H)))/(b_mz(c9)+(1*H))*1000000),4))&Chr(13)                      
										percent(d1)=1
										b_location(d1)=1
									Else
									End If
								else              
								end if              
								if oligofragments(c10)<CDbl((c_mz(c9)+(1*H)))+CDbl((c_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((c_mz(c9)+(1*H)))-CDbl((c_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                     
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										'MsgBox "m/z="&CStr(oligofragments(lightest_cluster_check_counter))&Chr(13)&"light_upper_window="&CStr(light_upper_window)&Chr(13)&"c_mz="&CStr((c_mz(c9)+(1*H)))&Chr(13)&"light_lower_window="&CStr(light_lower_window)
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End if
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)             
										f6.Write "c"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((c_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(cC(c9))+"H"+CStr(cH(c9)+1)+"F"+CStr(cF(c9))+"N"+CStr(cN(c9))+"O"+CStr(cO(c9))+"P"+CStr(cP(c9))+"S"+CStr(cS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(c_mz(c9)+(1*H)))/(c_mz(c9)+(1*H))*1000000),4))&Chr(13)              
										percent(d1)=1
										c_location(d1)=1
									Else
									End if
								else              
								end if                
								if oligofragments(c10)<CDbl((d_mz(c9)+(1*H)))+CDbl((d_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((d_mz(c9)+(1*H)))-CDbl((d_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then     
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next	
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)       
										f6.Write "d"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((d_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(dC(c9))+"H"+CStr(dH(c9)+1)+"F"+CStr(dF(c9))+"N"+CStr(dN(c9))+"O"+CStr(d_O(c9))+"P"+CStr(dP(c9))+"S"+CStr(dS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(d_mz(c9)+(1*H)))/(d_mz(c9)+(1*H))*1000000),4))&Chr(13)                     
										percent(d1)=1
										d_location(d1)=1
									Else
									end If
								else              
								end if              
								if oligofragments(c10)<CDbl((w_mz(c9)+(1*H)))+CDbl((w_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((w_mz(c9)+(1*H)))-CDbl((w_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then                 
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "w"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((w_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(wC(c9))+"H"+CStr(wH(c9)+1)+"F"+CStr(wF(c9))+"N"+CStr(wN(c9))+"O"+CStr(wO(c9))+"P"+CStr(wP(c9))+"S"+CStr(wS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(w_mz(c9)+(1*H)))/(w_mz(c9)+(1*H))*1000000),4))&Chr(13)              
										percent(c12-c9)=1
										w_location(c12-c9)=1
									Else
									End if
								else              
								end if              
								if oligofragments(c10)<CDbl((x_mz(c9)+(1*H)))+CDbl((x_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((x_mz(c9)+(1*H)))-CDbl((x_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then                    
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "x"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((x_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(xC(c9))+"H"+CStr(xH(c9)+1)+"F"+CStr(xF(c9))+"N"+CStr(xN(c9))+"O"+CStr(xO(c9))+"P"+CStr(xP(c9))+"S"+CStr(xS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(x_mz(c9)+(1*H)))/(x_mz(c9)+(1*H))*1000000),4))&Chr(13)                
										percent(c12-c9)=1
										x_location(c12-c9)=1
									Else
									end if
								else              
								end if                
								if oligofragments(c10)<CDbl((y_mz(c9)+(1*H)))+CDbl((y_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((y_mz(c9)+(1*H)))-CDbl((y_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then              
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
									'MsgBox "m/z="&CStr(oligofragments(lightest_cluster_check_counter))&Chr(13)&"light_upper_window="&CStr(light_upper_window)&Chr(13)&"y_mz="&CStr((y_mz(c9)+(1*H)))&Chr(13)&"light_lower_window="&CStr(light_lower_window)
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "y"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((y_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(yC(c9))+"H"+CStr(yH(c9)+1)+"F"+CStr(yF(c9))+"N"+CStr(yN(c9))+"O"+CStr(yO(c9))+"P"+CStr(yP(c9))+"S"+CStr(yS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(y_mz(c9)+(1*H)))/(y_mz(c9)+(1*H))*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										y_location(c12-c9)=1
									Else
									End if
								else      								
								end if                
								if oligofragments(c10)<CDbl((z_mz(c9)+(1*H)))+CDbl((z_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((z_mz(c9)+(1*H)))-CDbl((z_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then            
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "z"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((z_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(zC(c9))+"H"+CStr(zH(c9)+1)+"F"+CStr(zF(c9))+"N"+CStr(zN(c9))+"O"+CStr(zO(c9))+"P"+CStr(zP(c9))+"S"+CStr(zS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(z_mz(c9)+(1*H)))/(z_mz(c9)+(1*H))*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										z_location(c12-c9)=1
									else
									End If
								else              
								end if    
								'Next
								'c=c+1      
								c10=c10+1      
							Next 
						Else
							c10=0      
							For Each signal In oligofragments 
								For z=1 to max_charge_fragment
									if oligofragments(c10)<CDbl((a1_mz(c9)-(z-1)*H)/z)+CDbl((a1_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((a1_mz(c9)-(z-1)*H)/z)-CDbl((a1_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 then                      
										'annotate here       
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "a"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((a1_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(a1C(c9))+"H"+CStr(a1H(c9)-(z-1))+"F"+CStr(a1F(c9))+"N"+CStr(a1N(c9))+"O"+CStr(a1O(c9))+"P"+CStr(a1P(c9))+"S"+CStr(a1S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((a1_mz(c9)-(z-1)*H)/z))/((a1_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                        
										percent(d1)=1
										a_location(d1)=1
									else              
									End if              
									if oligofragments(c10)<CDbl((ab_mz(c9)-(z-1)*H)/z)+CDbl((ab_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((ab_mz(c9)-(z-1)*H)/z)-CDbl((ab_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 then                  
										'Cmpd(2).Annotations.AddAnnotation ab_mz_L, prime_int*1.2, ab_mz_H, prime_int*1.2, prime_int*1.6, "a-b"+CStr(d), True              
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "ab"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((ab_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(abC(c9))+"H"+CStr(abH(c9)-(z-1))+"F"+CStr(abF(c9))+"N"+CStr(abN(c9))+"O"+CStr(abO(c9))+"P"+CStr(abP(c9))+"S"+CStr(ab_S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((ab_mz(c9)-(z-1)*H)/z))/((ab_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)              
										percent(d1)=1
										ab_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((b_mz(c9)-(z-1)*H)/z)+CDbl((b_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((b_mz(c9)-(z-1)*H)/z)-CDbl((b_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 then                   
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "b"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((b_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(bC(c9))+"H"+CStr(bH(c9)-(z-1))+"F"+CStr(bF(c9))+"N"+CStr(bN(c9))+"O"+CStr(bO(c9))+"P"+CStr(bP(c9))+"S"+CStr(bS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((b_mz(c9)-(z-1)*H)/z))/((b_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                      
										percent(d1)=1
										b_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((c_mz(c9)-(z-1)*H)/z)+CDbl((c_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((c_mz(c9)-(z-1)*H)/z)-CDbl((c_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 then                     
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)             
										f6.Write "c"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((c_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(cC(c9))+"H"+CStr(cH(c9)-(z-1))+"F"+CStr(cF(c9))+"N"+CStr(cN(c9))+"O"+CStr(cO(c9))+"P"+CStr(cP(c9))+"S"+CStr(cS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((c_mz(c9)-(z-1)*H)/z))/(c_mz(c9))*1000000),4))&Chr(13)              
										percent(d1)=1
										c_location(d1)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((d_mz(c9)-(z-1)*H)/z)+CDbl((d_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((d_mz(c9)-(z-1)*H)/z)-CDbl((d_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<c12 then                   
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)       
										f6.Write "d"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((d_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(dC(c9))+"H"+CStr(dH(c9)-(z-1))+"F"+CStr(dF(c9))+"N"+CStr(dN(c9))+"O"+CStr(d_O(c9))+"P"+CStr(dP(c9))+"S"+CStr(dS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((d_mz(c9)-(z-1)*H)/z))/((d_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                     
										percent(d1)=1
										d_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((w_mz(c9)-(z-1)*H)/z)+CDbl((w_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((w_mz(c9)-(z-1)*H)/z)-CDbl((w_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 then                 
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "w"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((w_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(wC(c9))+"H"+CStr(wH(c9)-(z-1))+"F"+CStr(wF(c9))+"N"+CStr(wN(c9))+"O"+CStr(wO(c9))+"P"+CStr(wP(c9))+"S"+CStr(wS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((w_mz(c9)-(z-1)*H)/z))/((w_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)              
										'MsgBox "c12="&c12
										'MsgBox "c9="&c9
										percent(c12-c9)=1
										w_location(c12-c9)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((x_mz(c9)-(z-1)*H)/z)+CDbl((x_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((x_mz(c9)-(z-1)*H)/z)-CDbl((x_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 then                    
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "x"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((x_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(xC(c9))+"H"+CStr(xH(c9)-(z-1))+"F"+CStr(xF(c9))+"N"+CStr(xN(c9))+"O"+CStr(xO(c9))+"P"+CStr(xP(c9))+"S"+CStr(xS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((x_mz(c9)-(z-1)*H)/z))/((x_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                
										percent(c12-c9)=1
										x_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((y_mz(c9)-(z-1)*H)/z)+CDbl((y_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((y_mz(c9)-(z-1)*H)/z)-CDbl((y_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 then              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "y"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((y_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(yC(c9))+"H"+CStr(yH(c9)-(z-1))+"F"+CStr(yF(c9))+"N"+CStr(yN(c9))+"O"+CStr(yO(c9))+"P"+CStr(yP(c9))+"S"+CStr(yS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((y_mz(c9)-(z-1)*H)/z))/((y_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										y_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((z_mz(c9)-(z-1)*H)/z)+CDbl((z_mz(c9)-(z-1)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((z_mz(c9)-(z-1)*H)/z)-CDbl((z_mz(c9)-(z-1)*H)/z)*ppmrange2 and d1<=c12 then            
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "z"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(z)+Chr(44)+Chr(9)+CStr((z_mz(c9)-(z-1)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(zC(c9))+"H"+CStr(zH(c9)-(z-1))+"F"+CStr(zF(c9))+"N"+CStr(zN(c9))+"O"+CStr(zO(c9))+"P"+CStr(zP(c9))+"S"+CStr(zS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((z_mz(c9)-(z-1)*H)/z))/((z_mz(c9)-(z-1)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										z_location(c12-c9)=1
									else              
									end if    
								Next
								'c=c+1      
								c10=c10+1      
							Next 
						End if
                        c9=c9+1      
                    Next  
					c11=0
					sum=0
					'coverage=0
					For Each location in percent
						sum=sum+percent(c11)
						c11=c11+1
					Next
					'sequence coverage table writing
					'
					''
					f6.Write "ab1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(ab_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "a1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(a_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "b1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(b_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "c1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(c_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "d1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(d_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					'''
					''
					f6.Write "w"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(w_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "x"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(x_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "y"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(y_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "z"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(z_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					'
					f6.Write "percent coverage(%)="+CStr(sum/(c11-1)*100)&Chr(13)
                else              
                    'f6.Write oligoprecursor&Chr(13)            
                End If  
				c11=0
				For Each location in percent
					percent(c11)=0
					c11=c11+1
				Next
				'redefine boolean array for semi-visual sequence mapping
				'fill the boolean array with zeros by default
				c14=0
				For Each location in ab_location
					ab_location(c14)=0
					a_location(c14)=0
					b_location(c14)=0
					c_location(c14)=0
					d_location(c14)=0
					w_location(c14)=0
					x_location(c14)=0
					y_location(c14)=0
					z_location(c14)=0
					c14=c14+1
				Next
				c14=0
                'MsgBox negchargestates(c5)          
                'MsgBox oligoprecursor oligoprecursor         
                c5=c5+1              
            Next            
            c5=0                   
            For precursor=0 to max_charge_index        
                oligoprecursor_H=CDbl(poschargestates(c5))+CDbl(ppmrange2_precursor)              
                oligoprecursor_L=CDbl(poschargestates(c5))-CDbl(ppmrange2_precursor)      
                if    oligoprecursor<oligoprecursor_H and oligoprecursor>oligoprecursor_L then          
                    f6.Write CStr(tempfile(c8))&Chr(13)         
                    f6.Write "pos"+Chr(44)+Chr(9)+"exp="+CStr(oligoprecursor)+Chr(44)+Chr(9)+"theo="+CStr(poschargestates(c5))+Chr(44)+Chr(9)+oligoname(ccc-1)+Chr(44)+Chr(9)+"charge="+CStr(precursor+1)&Chr(13)      
                    c9=0      
                    d1=0    
					c12=0
					For Each location in percent
						c12=c12+1
					Next
                    For Each sugar In o_array1              
                        d1=c9+1              
                        'a1_mz_n(c)=CDbl(a1_mz(c))              
                        'a1_mz_L     =CDbl(a1_mz(c9)+2*H)-CDbl(a1_mz(c9)+2*H)*ppmrange2              
                        'a1_mz_H     =CDbl(a1_mz(c9)+2*H)+CDbl(a1_mz(c9)+2*H)*ppmrange2              
                        'ab_mz_L     =CDbl(ab_mz(c9)+2*H)-CDbl(ab_mz(c9)+2*H)*ppmrange2              
                        'ab_mz_H     =CDbl(ab_mz(c9)+2*H)+CDbl(ab_mz(c9)+2*H)*ppmrange2              
                        'b_mz_L     =CDbl(b_mz(c9)+2*H)-CDbl(b_mz(c9)+2*H)*ppmrange2              
                        'b_mz_H     =CDbl(b_mz(c9)+2*H)+CDbl(b_mz(c9)+2*H)*ppmrange2              
                        'c_mz_L     =CDbl(c_mz(c9)+2*H)-CDbl(c_mz(c9)+2*H)*ppmrange2              
                        'c_mz_H     =CDbl(c_mz(c9)+2*H)+CDbl(c_mz(c9)+2*H)*ppmrange2              
                        'd_mz_L     =CDbl(d_mz(c9)+2*H)-CDbl(d_mz(c9)+2*H)*ppmrange2              
                        'd_mz_H     =CDbl(d_mz(c9)+2*H)+CDbl(d_mz(c9)+2*H)*ppmrange2              
                        'w_mz_L     =CDbl(w_mz(c9)+2*H)-CDbl(w_mz(c9)+2*H)*ppmrange2              
                        'w_mz_H     =CDbl(w_mz(c9)+2*H)+CDbl(w_mz(c9)+2*H)*ppmrange2              
                        'x_mz_L     =CDbl(x_mz(c9)+2*H)-CDbl(x_mz(c9)+2*H)*ppmrange2              
                        'x_mz_H     =CDbl(x_mz(c9)+2*H)+CDbl(x_mz(c9)+2*H)*ppmrange2              
                        'y_mz_L     =CDbl(y_mz(c9)+2*H)-CDbl(y_mz(c9)+2*H)*ppmrange2              
                        'y_mz_H     =CDbl(y_mz(c9)+2*H)+CDbl(y_mz(c9)+2*H)*ppmrange2              
                        'z_mz_L     =CDbl(z_mz(c9)+2*H)-CDbl(z_mz(c9)+2*H)*ppmrange2              
                        'z_mz_H     =CDbl(z_mz(c9)+2*H)+CDbl(z_mz(c9)+2*H)*ppmrange2              
                        'MsgBox a1_mz_n_L           
                        'this is probably the place to put annother for loop to itterate through the fragmtn ion signals      
                        if charge_comparison="y" then
							c10=0      				
							For Each signal In oligofragments  
								
								For z=1 to max_charge_fragment
									if oligofragments(c10)<CDbl((a1_mz(c9)+(1+z)*H)/z)+CDbl((a1_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((a1_mz(c9)+(1+z)*H)/z)-CDbl((a1_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                      
										'annotate here       
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "a"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((a1_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(a1C(c9))+"H"+CStr(a1H(c9)+(z+1))+"F"+CStr(a1F(c9))+"N"+CStr(a1N(c9))+"O"+CStr(a1O(c9))+"P"+CStr(a1P(c9))+"S"+CStr(a1S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((a1_mz(c9)+(1+z)*H)/z))/((a1_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                        
										percent(d1)=1
										a_location(d1)=1
									else              
									End if              
									if oligofragments(c10)<CDbl((ab_mz(c9)+(1+z)*H)/z)+CDbl((ab_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((ab_mz(c9)+(1+z)*H)/z)-CDbl((ab_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                  
										'Cmpd(2).Annotations.AddAnnotation ab_mz_L, prime_int*1.2, ab_mz_H, prime_int*1.2, prime_int*1.6, "a-b"+CStr(d), True              
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "ab"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((ab_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(abC(c9))+"H"+CStr(abH(c9)+(z+1))+"F"+CStr(abF(c9))+"N"+CStr(abN(c9))+"O"+CStr(abO(c9))+"P"+CStr(abP(c9))+"S"+CStr(ab_S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((ab_mz(c9)+(1+z)*H)/z))/((ab_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)              
										percent(d1)=1
										ab_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((b_mz(c9)+(1+z)*H)/z)+CDbl((b_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((b_mz(c9)+(1+z)*H)/z)-CDbl((b_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                   
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "b"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((b_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(bC(c9))+"H"+CStr(bH(c9)+(z+1))+"F"+CStr(bF(c9))+"N"+CStr(bN(c9))+"O"+CStr(bO(c9))+"P"+CStr(bP(c9))+"S"+CStr(bS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((b_mz(c9)+(1+z)*H)/z))/((b_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                      
										percent(d1)=1
										b_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((c_mz(c9)+(1+z)*H)/z)+CDbl((c_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((c_mz(c9)+(1+z)*H)/z)-CDbl((c_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                     
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)             
										f6.Write "c"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((c_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(cC(c9))+"H"+CStr(cH(c9)+(z+1))+"F"+CStr(cF(c9))+"N"+CStr(cN(c9))+"O"+CStr(cO(c9))+"P"+CStr(cP(c9))+"S"+CStr(cS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((c_mz(c9)+(1+z)*H)/z))/((c_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)              
										percent(d1)=1
										c_location(d1)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((d_mz(c9)+(1+z)*H)/z)+CDbl((d_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((d_mz(c9)+(1+z)*H)/z)-CDbl((d_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 and CDbl(oligofragments_zs(c10))=z then                   
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)       
										f6.Write "d"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((d_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(dC(c9))+"H"+CStr(dH(c9)+2)+"F"+CStr(dF(c9))+"N"+CStr(dN(c9))+"O"+CStr(d_O(c9))+"P"+CStr(dP(c9))+"S"+CStr(dS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((d_mz(c9)+(1+z)*H)/z))/((d_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                     
										percent(d1)=1
										d_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((w_mz(c9)+(1+z)*H)/z)+CDbl((w_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((w_mz(c9)+(1+z)*H)/z)-CDbl((w_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then                 
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "w"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((w_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(wC(c9))+"H"+CStr(wH(c9)+(z+1))+"F"+CStr(wF(c9))+"N"+CStr(wN(c9))+"O"+CStr(wO(c9))+"P"+CStr(wP(c9))+"S"+CStr(wS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((w_mz(c9)+(1+z)*H)/z))/((w_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)              
										'MsgBox "c12="&c12
										'MsgBox "c9="&c9
										percent(c12-c9)=1
										w_location(c12-c9)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((x_mz(c9)+(1+z)*H)/z)+CDbl((x_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((x_mz(c9)+(1+z)*H)/z)-CDbl((x_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then                    
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "x"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((x_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(xC(c9))+"H"+CStr(xH(c9)+(z+1))+"F"+CStr(xF(c9))+"N"+CStr(xN(c9))+"O"+CStr(xO(c9))+"P"+CStr(xP(c9))+"S"+CStr(xS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((x_mz(c9)+(1+z)*H)/z))/((x_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                
										percent(c12-c9)=1
										x_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((y_mz(c9)+(1+z)*H)/z)+CDbl((y_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((y_mz(c9)+(1+z)*H)/z)-CDbl((y_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "y"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((y_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(yC(c9))+"H"+CStr(yH(c9)+(z+1))+"F"+CStr(yF(c9))+"N"+CStr(yN(c9))+"O"+CStr(yO(c9))+"P"+CStr(yP(c9))+"S"+CStr(yS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((y_mz(c9)+(1+z)*H)/z))/((y_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										y_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((z_mz(c9)+(1+z)*H)/z)+CDbl((z_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((z_mz(c9)+(1+z)*H)/z)-CDbl((z_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 and CDbl(oligofragments_zs(c10))=z then            
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "z"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((z_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(zC(c9))+"H"+CStr(zH(c9)+(z+1))+"F"+CStr(zF(c9))+"N"+CStr(zN(c9))+"O"+CStr(zO(c9))+"P"+CStr(zP(c9))+"S"+CStr(zS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((z_mz(c9)+(1+z)*H)/z))/((z_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										z_location(c12-c9)=1
									else              
									end if     
								Next
								'c=c+1      
								c10=c10+1      
							Next   
						Elseif charge_comparison="x" then
							c10=0      
							For Each signal In oligofragments 
								'For z=1 to max_charge_fragment
								if oligofragments(c10)<CDbl((a1_mz(c9)+(1*H)))+CDbl((a1_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((a1_mz(c9)+(1*H)))-CDbl((a1_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                      
									'the second if statmetn below is to check if the signal in the oligofragments has a signal which is between 0.97 and 1.03 daltons lower mass than the signal itself, if false this peak is not a descernable monooisotopic peak in this workflow
									'else if the peak is a descernable monoisotopic peak it may be assigned
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "a"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((a1_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(a1C(c9))+"H"+CStr(a1H(c9)+1)+"F"+CStr(a1F(c9))+"N"+CStr(a1N(c9))+"O"+CStr(a1O(c9))+"P"+CStr(a1P(c9))+"S"+CStr(a1S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(a1_mz(c9)+(1*H)))/(a1_mz(c9)+(1*H))*1000000),4))&Chr(13)                        
										percent(d1)=1
										a_location(d1)=1
									Else
									End if
								else              
								End if              
								if oligofragments(c10)<CDbl((ab_mz(c9)+(1*H)))+CDbl((ab_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((ab_mz(c9)+(1*H)))-CDbl((ab_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                  
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "ab"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((ab_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(abC(c9))+"H"+CStr(abH(c9)+1)+"F"+CStr(abF(c9))+"N"+CStr(abN(c9))+"O"+CStr(abO(c9))+"P"+CStr(abP(c9))+"S"+CStr(ab_S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(ab_mz(c9)+(1*H)))/(ab_mz(c9)+(1*H))*1000000),4))&Chr(13)              
										percent(d1)=1
										ab_location(d1)=1
									Else
									End if
								else              
								end if              
								if oligofragments(c10)<CDbl((b_mz(c9)+(1*H)))+CDbl((b_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((b_mz(c9)+(1*H)))-CDbl((b_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                   
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next	
									if light_logic=true then
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "b"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((b_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(bC(c9))+"H"+CStr(bH(c9)+1)+"F"+CStr(bF(c9))+"N"+CStr(bN(c9))+"O"+CStr(bO(c9))+"P"+CStr(bP(c9))+"S"+CStr(bS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(b_mz(c9)+(1*H)))/(b_mz(c9)+(1*H))*1000000),4))&Chr(13)                      
										percent(d1)=1
										b_location(d1)=1
									Else
									End If
								else              
								end if              
								if oligofragments(c10)<CDbl((c_mz(c9)+(1*H)))+CDbl((c_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((c_mz(c9)+(1*H)))-CDbl((c_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then                     
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										'MsgBox "m/z="&CStr(oligofragments(lightest_cluster_check_counter))&Chr(13)&"light_upper_window="&CStr(light_upper_window)&Chr(13)&"c_mz="&CStr((c_mz(c9)+(1*H)))&Chr(13)&"light_lower_window="&CStr(light_lower_window)
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End if
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)             
										f6.Write "c"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((c_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(cC(c9))+"H"+CStr(cH(c9)+1)+"F"+CStr(cF(c9))+"N"+CStr(cN(c9))+"O"+CStr(cO(c9))+"P"+CStr(cP(c9))+"S"+CStr(cS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(c_mz(c9)+(1*H)))/(c_mz(c9)+(1*H))*1000000),4))&Chr(13)              
										percent(d1)=1
										c_location(d1)=1
									Else
									End if
								else              
								end if                
								if oligofragments(c10)<CDbl((d_mz(c9)+(1*H)))+CDbl((d_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((d_mz(c9)+(1*H)))-CDbl((d_mz(c9)+(1*H)))*ppmrange2 and d1<c12 then     
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next	
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)       
										f6.Write "d"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((d_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(dC(c9))+"H"+CStr(dH(c9)+1)+"F"+CStr(dF(c9))+"N"+CStr(dN(c9))+"O"+CStr(d_O(c9))+"P"+CStr(dP(c9))+"S"+CStr(dS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(d_mz(c9)+(1*H)))/(d_mz(c9)+(1*H))*1000000),4))&Chr(13)                     
										percent(d1)=1
										d_location(d1)=1
									Else
									end If
								else              
								end if              
								if oligofragments(c10)<CDbl((w_mz(c9)+(1*H)))+CDbl((w_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((w_mz(c9)+(1*H)))-CDbl((w_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then                 
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "w"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((w_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(wC(c9))+"H"+CStr(wH(c9)+1)+"F"+CStr(wF(c9))+"N"+CStr(wN(c9))+"O"+CStr(wO(c9))+"P"+CStr(wP(c9))+"S"+CStr(wS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(w_mz(c9)+(1*H)))/(w_mz(c9)+(1*H))*1000000),4))&Chr(13)              
										percent(c12-c9)=1
										w_location(c12-c9)=1
									Else
									End if
								else              
								end if              
								if oligofragments(c10)<CDbl((x_mz(c9)+(1*H)))+CDbl((x_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((x_mz(c9)+(1*H)))-CDbl((x_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then                    
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "x"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((x_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(xC(c9))+"H"+CStr(xH(c9)+1)+"F"+CStr(xF(c9))+"N"+CStr(xN(c9))+"O"+CStr(xO(c9))+"P"+CStr(xP(c9))+"S"+CStr(xS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(x_mz(c9)+(1*H)))/(x_mz(c9)+(1*H))*1000000),4))&Chr(13)                
										percent(c12-c9)=1
										x_location(c12-c9)=1
									Else
									end if
								else              
								end if                
								if oligofragments(c10)<CDbl((y_mz(c9)+(1*H)))+CDbl((y_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((y_mz(c9)+(1*H)))-CDbl((y_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then              
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "y"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((y_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(yC(c9))+"H"+CStr(yH(c9)+1)+"F"+CStr(yF(c9))+"N"+CStr(yN(c9))+"O"+CStr(yO(c9))+"P"+CStr(yP(c9))+"S"+CStr(yS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(y_mz(c9)+(1*H)))/(y_mz(c9)+(1*H))*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										y_location(c12-c9)=1
									Else
									End if
								else      								
								end if                
								if oligofragments(c10)<CDbl((z_mz(c9)+(1*H)))+CDbl((z_mz(c9)+(1*H)))*ppmrange2 and oligofragments(c10)>CDbl((z_mz(c9)+(1*H)))-CDbl((z_mz(c9)+(1*H)))*ppmrange2 and d1<=c12 then            
									light_logic=true
									light_upper_window=oligofragments(c10)-small_dalton
									light_lower_window=oligofragments(c10)-big_dalton
									for lightest_cluster_check_counter=1 to oligofragmentcount
										light_check=oligofragments(lightest_cluster_check_counter)
										if light_check<light_upper_window and light_check>light_lower_window then
											light_logic=false
										else
										End If
									Next
									if light_logic=true then
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "z"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(45)+CStr(0)+Chr(44)+Chr(9)+CStr((z_mz(c9)+(1*H)))+Chr(44)+Chr(9)+"C"+CStr(zC(c9))+"H"+CStr(zH(c9)+1)+"F"+CStr(zF(c9))+"N"+CStr(zN(c9))+"O"+CStr(zO(c9))+"P"+CStr(zP(c9))+"S"+CStr(zS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-(z_mz(c9)+(1*H)))/(z_mz(c9)+(1*H))*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										z_location(c12-c9)=1
									else
									End If
								else              
								end if    
								'Next
								'c=c+1      
								c10=c10+1      
							Next 
						Else
							c10=0      				
							For Each signal In oligofragments  
								
								For z=1 to max_charge_fragment
									if oligofragments(c10)<CDbl((a1_mz(c9)+(1+z)*H)/z)+CDbl((a1_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((a1_mz(c9)+(1+z)*H)/z)-CDbl((a1_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 then                      
										'annotate here       
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "a"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((a1_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(a1C(c9))+"H"+CStr(a1H(c9)+(z+1))+"F"+CStr(a1F(c9))+"N"+CStr(a1N(c9))+"O"+CStr(a1O(c9))+"P"+CStr(a1P(c9))+"S"+CStr(a1S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((a1_mz(c9)+(1+z)*H)/z))/((a1_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                        
										percent(d1)=1
										a_location(d1)=1
									else              
									End if              
									if oligofragments(c10)<CDbl((ab_mz(c9)+(1+z)*H)/z)+CDbl((ab_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((ab_mz(c9)+(1+z)*H)/z)-CDbl((ab_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 then                  
										'Cmpd(2).Annotations.AddAnnotation ab_mz_L, prime_int*1.2, ab_mz_H, prime_int*1.2, prime_int*1.6, "a-b"+CStr(d), True              
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "ab"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((ab_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(abC(c9))+"H"+CStr(abH(c9)+(z+1))+"F"+CStr(abF(c9))+"N"+CStr(abN(c9))+"O"+CStr(abO(c9))+"P"+CStr(abP(c9))+"S"+CStr(ab_S(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((ab_mz(c9)+(1+z)*H)/z))/((ab_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)              
										percent(d1)=1
										ab_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((b_mz(c9)+(1+z)*H)/z)+CDbl((b_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((b_mz(c9)+(1+z)*H)/z)-CDbl((b_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 then                   
										'annotate here              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "b"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((b_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(bC(c9))+"H"+CStr(bH(c9)+(z+1))+"F"+CStr(bF(c9))+"N"+CStr(bN(c9))+"O"+CStr(bO(c9))+"P"+CStr(bP(c9))+"S"+CStr(bS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((b_mz(c9)+(1+z)*H)/z))/((b_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                      
										percent(d1)=1
										b_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((c_mz(c9)+(1+z)*H)/z)+CDbl((c_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((c_mz(c9)+(1+z)*H)/z)-CDbl((c_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 then                     
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)             
										f6.Write "c"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((c_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(cC(c9))+"H"+CStr(cH(c9)+(z+1))+"F"+CStr(cF(c9))+"N"+CStr(cN(c9))+"O"+CStr(cO(c9))+"P"+CStr(cP(c9))+"S"+CStr(cS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((c_mz(c9)+(1+z)*H)/z))/((c_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)              
										percent(d1)=1
										c_location(d1)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((d_mz(c9)+(1+z)*H)/z)+CDbl((d_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((d_mz(c9)+(1+z)*H)/z)-CDbl((d_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<c12 then                   
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)       
										f6.Write "d"+Chr(44)+Chr(9)+CStr(d1)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((d_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(dC(c9))+"H"+CStr(dH(c9)+2)+"F"+CStr(dF(c9))+"N"+CStr(dN(c9))+"O"+CStr(d_O(c9))+"P"+CStr(dP(c9))+"S"+CStr(dS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((d_mz(c9)+(1+z)*H)/z))/((d_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                     
										percent(d1)=1
										d_location(d1)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((w_mz(c9)+(1+z)*H)/z)+CDbl((w_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((w_mz(c9)+(1+z)*H)/z)-CDbl((w_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 then                 
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "w"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((w_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(wC(c9))+"H"+CStr(wH(c9)+(z+1))+"F"+CStr(wF(c9))+"N"+CStr(wN(c9))+"O"+CStr(wO(c9))+"P"+CStr(wP(c9))+"S"+CStr(wS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((w_mz(c9)+(1+z)*H)/z))/((w_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)              
										'MsgBox "c12="&c12
										'MsgBox "c9="&c9
										percent(c12-c9)=1
										w_location(c12-c9)=1
									else              
									end if              
									if oligofragments(c10)<CDbl((x_mz(c9)+(1+z)*H)/z)+CDbl((x_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((x_mz(c9)+(1+z)*H)/z)-CDbl((x_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 then                    
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "x"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((x_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(xC(c9))+"H"+CStr(xH(c9)+(z+1))+"F"+CStr(xF(c9))+"N"+CStr(xN(c9))+"O"+CStr(xO(c9))+"P"+CStr(xP(c9))+"S"+CStr(xS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((x_mz(c9)+(1+z)*H)/z))/((x_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                
										percent(c12-c9)=1
										x_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((y_mz(c9)+(1+z)*H)/z)+CDbl((y_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((y_mz(c9)+(1+z)*H)/z)-CDbl((y_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 then              
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "y"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((y_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(yC(c9))+"H"+CStr(yH(c9)+(z+1))+"F"+CStr(yF(c9))+"N"+CStr(yN(c9))+"O"+CStr(yO(c9))+"P"+CStr(yP(c9))+"S"+CStr(yS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((y_mz(c9)+(1+z)*H)/z))/((y_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										y_location(c12-c9)=1
									else              
									end if                
									if oligofragments(c10)<CDbl((z_mz(c9)+(1+z)*H)/z)+CDbl((z_mz(c9)+(1+z)*H)/z)*ppmrange2 and oligofragments(c10)>CDbl((z_mz(c9)+(1+z)*H)/z)-CDbl((z_mz(c9)+(1+z)*H)/z)*ppmrange2 and d1<=c12 then            
										f6.Write CStr(oligofragments(c10))+Chr(44)+Chr(9)      
										f6.Write "z"+Chr(44)+Chr(9)+CStr(c9)+Chr(44)+Chr(9)+Chr(43)+CStr(z)+Chr(44)+Chr(9)+CStr((z_mz(c9)+(1+z)*H)/z)+Chr(44)+Chr(9)+"C"+CStr(zC(c9))+"H"+CStr(zH(c9)+(z+1))+"F"+CStr(zF(c9))+"N"+CStr(zN(c9))+"O"+CStr(zO(c9))+"P"+CStr(zP(c9))+"S"+CStr(zS(c9))+Chr(44)+Chr(9)+CStr(Round(((oligofragments(c10)-((z_mz(c9)+(1+z)*H)/z))/((z_mz(c9)+(1+z)*H)/z)*1000000),4))&Chr(13)                     
										percent(c12-c9)=1
										z_location(c12-c9)=1
									else              
									end if     
								Next
								'c=c+1      
								c10=c10+1      
							Next   
						End if
                        c9=c9+1      
                    Next  
					c11=0
					sum=0
					'coverage=0
					For Each location in percent
						sum=sum+percent(c11)
						c11=c11+1
					Next
					
					'sequence coverage table writing
					'
					
					''
					f6.Write "ab1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(ab_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "a1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(a_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "b1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(b_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "c1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(c_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "d1"+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(d_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					'''
					''
					f6.Write "w"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(w_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "x"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(x_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "y"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(y_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					''
					f6.Write "z"+CStr(c11-1)+Chr(9)+"-"
					c13=1
					For location=1 to (c11-1)
						f6.Write CStr(z_location(c13))+Chr(44)+Chr(32)
						c13=c13+1
					Next
					f6.Write Chr(13)
					''
					'
					
					f6.Write "percent coverage(%)="+CStr(sum/(c11-1)*100)&Chr(13)
                else           
                    'f6.Write oligoprecursor_H&Chr(13)            
                End If          
                c5=c5+1         
            Next     
			
            f6.Close   
        Next       
        'if c2=100 then              
        '    '----' 
		'	Dim WshShell               
		'	Set WshShell = WScript.CreateObject("WScript.Shell")  			
        '    popup = WshShell.Popup("MS/MS spectra read="&d&"/"&begin_ion_count, 0.5)               
        '    '----'              
        '    c2=0              
        'else              
        'End If              
        c2=c2+1              
        MyFile.Close   
		Set MyFile=Nothing
        c6=c6+1         
        d=d+1              
        c4=c4+1              
    Next              
'------------------------------------------------------------------                  
'the interface code whichsets the user options              
'--------------------------------------------              
    ccc=0                       
c8=c8+1             
'MsgBox c8         
Next
Set f6 = fso.OpenTextFile(fragfilepath+"\"+"report.txt", 8, True)
f6.Write CStr(WScript.ScriptName)&Chr(13)
f6.Write "user inputs:"&Chr(13)
f6.Write "precoursor mz tolerance="+CStr(ppmrange_precursor)&Chr(13)
f6.Write "fragment mz tolerance="+CStr(ppmrange)&Chr(13)
f6.Write "max charge of precoursor searched="+CStr(max_charge_per_base)&Chr(13)
f6.Write "max charge of fragment searched="+CStr(max_charge_fragment)&Chr(13)
f6.Write "require charge state comparison="+CStr(charge_comparison)&Chr(13)
c15=0	
f6.Write "sequences searched:"&Chr(13)	
For Each oligo in oligos              
    f6.Write CStr(oligos(c15))&Chr(13)     
    c15=c15+1              
Next  
			
f6.Close

MyTime2=Time
'MsgBox Mytime1
'MsgBox MyTime2

MsgBox CStr(Mytime1)+Chr(13)+CStr(Mytime2)+Chr(13)+"Done"
Set objshell  = Nothing
WScript.Quit()
WSH.Quit()





