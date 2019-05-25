public sub rebar()

Dim varPick As Variant
'integreate all this inside safeie not inside the loop

With ThisDrawing.Utility
    varPick = .GetPoint(, vbCr & "Pick a point: ")
    .Prompt vbCr & varPick(0) & "," & varPick(1)
	

	
' 1.pick a point.. var(pick).. the picked point has relation with
	' the beam_dwg with like the variable say one that is with rebar 
		' aix, asj, asy,or asi 
	' go a distance of 1695.1543... and locate ra(0)

dim ra(0 to 2) as double
ra(0)=asj(0)
ra(1)=asj(1)+1695.1543
Set line = ThisDrawing.ModelSpace.AddLine(ra, asj)
line.Update

' 2.with ra(0).. go a distance of (bw-2*cover) down..rax(0)
	' and add a curve that with degree of 45 with
	' distance of.. with rai(0) var

' going distance down
	dim disb(0 to 2) as double
	disb(0)=ra(0)
	disb(1)=ra(1)+350 'rbs=sec - (25+25)
	Set line = ThisDrawing.ModelSpace.AddLine(ra, disb)
	line.Update
' doing the curver
	dim dicu(0 to 2) as double
	dicu(0)=disb(0)+128.774
	dicu(1)=disb(1)+67.878 
	Set line = ThisDrawing.ModelSpace.AddLine(dicu, disb)
	line.Update


' 3. then with ra(0) go a distance say having a relation

identify the number of rebar...

	if nb >= 4 then
		' have here two reinforcement arrows for the top bar
		*span_1 + span_2...+ rbs < 12
		*then find the previous from the last span..
		*try doing the overlap.. here..dependent on dimeter of bar	
			*......part_1........
			*go at a distance of overlap/2.. and do the curve
			*at the half of the journey find the coordinate().. do the curve
			*go up at distance of 148.2 up and go a dist of overlap/2 back and coordinate()
			*do the curve then
			*......part_1_End........
			*from there	test again span_3 + span_4 + recentSpan/2..<12
				how to know the last section part?
			*go a distance down of 148.2 
				*repeat part_1...
			*
			*
			*
			*
			*
			*
		*
		*
		*
		
		
		
		
	else
		' proceed just like ordinary
		
		
		
	end if 
		

' 4. 


' 5. 

End With


end sub
