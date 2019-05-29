Attribute VB_Name = "Module11"
Public Sub TestGetPoint()
' ////////////////////////////////////////////////////////////////////////
' Around here creatte an enviroment that exist just like beam detail drawing

' 1.layer


' 2.text style

' //////////////////////////////////////////////////////////////////////////

Dim varPick As Variant

With ThisDrawing.Utility
    varPick = .GetPoint(, vbCr & "Pick a point: ")
    .Prompt vbCr & varPick(0) & "," & varPick(1)
         
        Dim o As Integer
        Dim sc As Double
        Dim sm As Double
        Dim m As Integer
        Dim n As Integer
        Dim g(0 To 2) As Double
        Dim u(0 To 2) As Double

'________________ SpanNumber_____________________
o = 5 'number of span... is 5
'________________//SpanNumber_____________________


Dim p(0 To 100) As Double
Dim f(0 To 100) As Double
Dim h(0 To 100) As Double
Dim s(0 To 100) As Double

'________________ spanDimension______________________

p(0) = 2700
p(1) = 2000

p(2) = 2500
p(3) = 2500
p(4) = 2500
'________________//spanDimension//______________________


'________________ columnDim______________________
 
f(0) = 400
f(1) = 500
f(2) = 600
f(3) = 800
f(4) = 200
f(5) = 400 ' this last number need to be equal with span_number
 
 
'________________ //columnDim// ______________________

'assigning new values to new variables
For pn = 1 To o
    h(0) = f(0)
    h(pn) = f(pn)
    s(0) = p(0)
    s(pn) = p(pn)
Next

For sn = 1 To o

    If sn = 1 Then
                    ' TheFirstPart
                    Dim x(0 To 2) As Double
                    x(0) = varPick(0)
                    x(1) = varPick(1)
                    
                    Dim y(0 To 2) As Double
                    y(0) = x(0)
                    y(1) = x(1) + 186.5455

                    Set line = ThisDrawing.ModelSpace.AddLine(x, y)

                    Dim z(0 To 2) As Double
                    z(0) = y(0) + p(sn - 1)
                    z(1) = x(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(x, z)
                    line.Update
                    
                    Dim dzi(0 To 2) As Double
                    dzi(0) = z(0)
                    dzi(1) = z(1) - 25
'                    Set line = ThisDrawing.ModelSpace.AddLine(dzi, z)
'                    line.Update
                    
                    ' loop for sheear rein
                    Dim nio As Integer
                    Dim nis As Integer
                    nio = 3
                    For nis = 1 To nio
                        Dim shr(0 To 2) As Double
                        shr(0) = dzi(0)
                        shr(1) = dzi(1) - 200 + 25 + 25
                        Set line = ThisDrawing.ModelSpace.AddLine(shr, dzi)
                        line.Update
                            
                        dzi(0) = dzi(0) - 120
                        dzi(1) = dzi(1)
                            
                    Next nis
                    
                    Dim i(0 To 2) As Double
                    i(0) = z(0)
                    i(1) = z(1) + 186.5455
                        
                    Set line = ThisDrawing.ModelSpace.AddLine(z, i)
                    
                    ' moving from x() 25m cover down..
                    Dim asi(0 To 2) As Double
                    asi(0) = x(0)
                    asi(1) = x(1) - 25 ' notice here 25 is cover lenght...

                    Dim aix(0 To 2) As Double
                    aix(0) = x(0)
                    aix(1) = x(1) - 25 ' notice here 25 is cover lenght...
                    

'                    Set line = ThisDrawing.ModelSpace.AddLine(asi, x)
'                    line.Update
'
                    Dim asj(0 To 2) As Double
                    asj(0) = x(0) - f(0) + 25
                    asj(1) = asi(1)
'                    Set line = ThisDrawing.ModelSpace.AddLine(asj, asi)
'                    line.Update
                    
                    
                    Dim asx(0 To 2) As Double
                    asx(0) = x(0)
                    asx(1) = x(1) - (200) + 25 'notice here 200 is beam width
'                    Set line = ThisDrawing.ModelSpace.AddLine(asx, asi)
'                    line.Update
                    
                    Dim asy(0 To 2) As Double
                    asy(0) = x(0) - f(0) + 25
                    asy(1) = asx(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(asy, asj)
                    line.Update
                    

                    nio = 3
                    For nis = 1 To nio
                        Dim ali(0 To 2) As Double
                        ali(0) = asi(0)
                        ali(1) = asi(1) - 200 + 25 + 25
                        Set line = ThisDrawing.ModelSpace.AddLine(asi, ali)
                        line.Update
                            
                        asi(0) = asi(0) + 120
                        asi(1) = asi(1)
                        
                    Next nis
                    
                    
                    ' here is where ... beam width is  added..b1 say
                    Dim b(0 To 2) As Double
                    b(0) = y(0)
                    b(1) = x(1) - 200 '.....the length

                    'add new variable for dim.. jis()
                    Dim jis(0 To 2) As Double
                    jis(0) = b(0)
                    jis(1) = b(1) + (200 / 2)
                     ' Set line = ThisDrawing.ModelSpace.AddLine(jis, b)
'                     line.Update

                    'add new variable for dim.. jii()
                    Dim jii(0 To 2) As Double
                    jii(0) = z(0)
                    jii(1) = z(1) - (200 / 2)
                     ' Set line = ThisDrawing.ModelSpace.AddLine(jii, z)
'                     line.Update
                    ' add new dimension properties here for jis() & jii()..
                    Dim ptc1(0 To 2) As Double
                    Dim ptc2(0 To 2) As Double
                    Dim lod(0 To 2) As Double
                    
                    ptc1(0) = jis(0)
                    ptc1(1) = jis(1)

                    ptc2(0) = jii(0)
                    ptc2(1) = jii(1)

                    lod(0) = jis(0) / 2
                    lod(1) = ptc1(1)

                    rotAngle = 0
                    rotAngle = rotAngle * 3.141592 / 180#       ' covert to Radians
                                                                                    
                    'Add dimension
                    Set aDim = ThisDrawing.ModelSpace.AddDimRotated(ptc1(), ptc2(), lod(), rotAngle)

                    'Set dimension properties
                    aDim.color = acByLayer

                    'aDim.ExtensionLineExtend = 0

                    aDim.LinetypeScale = 100

                    aDim.Arrowhead1Type = acclosedfilled
                    aDim.Arrowhead2Type = acclosedfilled
                    '        aDim.arrowsize
                    aDim.ArrowheadSize = 100
                    aDim.TextColor = RGB(255, 127, 0)
                    ' aDim.TextColor = RGB(255, 127, 0)
                    'notice here 200 = beam width
                
                    aDim.TextHeight = 85
                    ' aDim.TextHeight = 220
                    aDim.UnitsFormat = acDimLDecimal

                    aDim.ExtLine1Suppress = True
                    aDim.ExtLine2Suppress = True

                    aDim.PrimaryUnitsPrecision = acDimPrecisionOne
                    aDim.TextGap = 30
                    ' aDim.TextGap = 3
                    aDim.LinearScaleFactor = 1
                    ' aDim.LinearScaleFactor = 1
                    aDim.ExtensionLineOffset = 0
                    ' aDim.ExtensionLineOffset = 1000
                    
                    aDim.VerticalTextPosition = acCentered
                    ' aDim.VerticalTextPosition = acAbove

                    aDim.PrimaryUnitsPrecision = acDimPrecisionZero
                    'Create a new dimension style
                    
                    aDim.TextOverride = "{\fUtsaah|b0|i0|c0|p34;8@c/c180}"
                    'aDim.TextStyle = sdf
                    ThisDrawing.ActiveTextStyle.height = 85
                    Set dimstyle = ThisDrawing.DimStyles.Add("D100")

                    'Create a new dimension style
                    'Set dimstyle = ThisDrawing.DimStyles.Add("jjkj")


                    aDim.Update
                    
                    Dim c(0 To 2) As Double
                    c(0) = y(0)
                    c(1) = b(1) - 186.5455

                '    _____186.545.. is the length b/n c-b...
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(c, b)
                    line.Update
                    Dim d(0 To 2) As Double
                    d(0) = b(0) + p(sn - 1) ' here is where l(1) goes
                    d(1) = b(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(b, d)
                    line.Update
                    Dim e(0 To 2) As Double
                    e(0) = d(0)
                    e(1) = d(1) - 186.5455
                
                    Set line = ThisDrawing.ModelSpace.AddLine(d, e)
                    line.Update
                 
                    ' bottom
                    
                    Dim k(0 To 2) As Double
                    k(0) = c(0) - f(0) ' here add column width c1
                    k(1) = c(1)

                    Dim kk(0 To 2) As Double
                    kk(0) = k(0) + ((f(0) / 2) - (73.418 / 2))
                    kk(1) = k(1)
                

                    Set line = ThisDrawing.ModelSpace.AddLine(kk, k)
                    line.Update
                    
                     Dim gg(0 To 2) As Double
                        gg(0) = kk(0) + 24.1052
                        gg(1) = kk(1) + 70.9835
                        
                    Set line = ThisDrawing.ModelSpace.AddLine(gg, kk)
                    line.Update
                    Dim cc(0 To 2) As Double
                    cc(0) = c(0) - ((f(0) / 2) - (73.418 / 2))
                    cc(1) = c(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(cc, c)
                    
                    Dim hh(0 To 2) As Double
                        hh(0) = cc(0) - 20.3504
                        hh(1) = cc(1) - 38.8248
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(hh, cc)
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(gg, hh)
                        
                        
                        line.Update
                    ' End_bottom zigzag
                          
                    ' top zigzag
                    Dim j(0 To 2) As Double
                    j(0) = x(0) - f(0) 'here add c1
                    j(1) = y(1)  '  jtotal = 2*186.5455 + b1 here add c1
                    Set line = ThisDrawing.ModelSpace.AddLine(k, j)
                    line.Update
                     Dim ee(0 To 2) As Double
                    ee(0) = j(0) + ((f(0) / 2) - (73.418 / 2))
                    ee(1) = j(1)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(ee, j)
                    line.Update
                    Dim aa(0 To 2) As Double
                        aa(0) = ee(0) + 24.1052
                        aa(1) = ee(1) + 70.9835
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(aa, ee)
                    
                     Dim ff(0 To 2) As Double
                    ff(0) = y(0) - ((f(0) / 2) - (73.418 / 2))
                    ff(1) = y(1)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(ff, y)

                    line.Update
                    
                     Dim jj(0 To 2) As Double
                    ' bottom constants
                    jj(0) = j(0) - 60
                    jj(1) = j(1)
                    
                     Dim yy(0 To 2) As Double
                    yy(0) = y(0) + 60
                    yy(1) = y(1)

                    'End bottom constants
                    Dim vv(0 To 2) As Double
                    ' top constants
                    vv(0) = k(0) - 60
                    vv(1) = k(1)
                    
                    Dim xx(0 To 2) As Double
                    xx(0) = c(0) + 60
                    xx(1) = c(1)
                    
                    'End top constants

                    Set line = ThisDrawing.ModelSpace.AddLine(jj, j)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(vv, k)
                    Set line = ThisDrawing.ModelSpace.AddLine(xx, c)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(yy, y)
                    line.Update
                     Dim bb(0 To 2) As Double
                        bb(0) = ff(0) - 20.3504
                        bb(1) = ff(1) - 38.8248
                        
                     Set line = ThisDrawing.ModelSpace.AddLine(ff, bb)

                     Set line = ThisDrawing.ModelSpace.AddLine(aa, bb)
                            line.Update
                            
                    ' END top zigzag
                            
                            
                    Dim tca(0 To 2) As Double
                    tca(0) = j(0) + (f(0) / 2)
                    tca(1) = j(1)

                    ' b straight line up..coodinate..
                    'to find b coordinate
                    Dim tcb(0 To 2) As Double
                    tcb(0) = tca(0)
                    tcb(1) = tca(1) + 335

                    Set line = ThisDrawing.ModelSpace.AddLine(tcb, tca)
                    
                    Dim tcc(0 To 2) As Double
                    ' c center line from  radius of circle find coodinate
                    tcc(0) = tcb(0)
                    tcc(1) = tcb(1) + 194
                    ' ... draw circle.. with the above coordinate....tcc(0) and tcc(1)

                    Dim objEnt As AcadCircle
                    Set objEnt = ThisDrawing.ModelSpace.AddCircle(tcc, 194)

                    Dim tcd(0 To 2) As Double
                    ' d coodinate frm b coodinate... add dia of circle
                    tcd(0) = tcc(0)
                    tcd(1) = tcc(1) + 194
                    
                    Dim tce(0 To 2) As Double
                    ' e coodinate frm d coodinate... add dia of dist_ed
                    tce(0) = tcd(0)
                    tce(1) = tcd(1) + 147

                    Set line = ThisDrawing.ModelSpace.AddLine(tcd, tce)
                    
                    Dim tcf(0 To 2) As Double
                    ' f coodinate frm c centerCircle..
                    tcf(0) = tcc(0) - 194
                    tcf(1) = tcc(1)
                    
                    Dim tch(0 To 2) As Double
                    ' h coordinate frm f coordinate add dist_fh
                    tch(0) = tcf(0) - 160
                    tch(1) = tcf(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(tch, tcf)
                    
                    Dim tcg(0 To 2) As Double
                    ' g coodinate frm c centerCircle ..
                    tcg(0) = tcc(0) + 194
                    tcg(1) = tcc(1)

                    ' i coodinate frm g coordinate add dist_gi
                    Dim tci(0 To 2) As Double
                    tci(0) = tcg(0) + 160
                    tci(1) = tcg(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(tci, tcg)
                    ' add text properties
                    
                    Dim textObj As AcadMText
                    Dim height As Double
                    height = 200
                    Dim tcj(0 To 2) As Double
                    tcj(0) = tcc(0) - 69.6153
                    tcj(1) = tcc(1) + 90.4984
                    
                    'what if the number is letter? block letter???
                    Set textObj = ThisDrawing.ModelSpace.AddMText(tcj, 200, sn)
                    textObj.height = 200
                        
                    line.Update
                    
                ' 'starts
                ' 1. with k(0).. coordinate.. find the cetner of the column cc(0)...using the f(0) column width

                    Dim ci(0 To 2) As Double
                    
                    cc(0) = k(0) + (f(0) / 2) ' here add column width c1
                    cc(1) = k(1)

                    ci(0) = j(0) + (f(0) / 2) ' here add column width c1
                    ci(1) = j(1)
                        
                ' 2.with dist s = 305  find coodinate... call it ...dio(0).... dio(1)..
                    Dim dis(0 To 2) As Double
                    dis(0) = cc(0) ' here add column width c1
                    dis(1) = cc(1) - 305
                
                ' some how get this coodinate then  find the next coodinate..
                ' to do for the inner.. one... and then finally.. proceed to the end...
                
                ' 3. add line... with dist = 500+305 fm center of column...coodinate...
                    Dim die(0 To 2) As Double
                    die(0) = dis(0) ' here add column width c1
                    die(1) = dis(1) - 500 'this is not column width..
                
                    Set line = ThisDrawing.ModelSpace.AddLine(die, ci)
                    line.Update







    Else
    
                    For m = 2 To sn
                       If m = 2 Then
                           sm = p(m - 2)
                       Else
                       sm = sm + p(m - 2)
                       End If
                    Next m
                    
                    For n = 2 To sn
                       If n = 2 Then
                           sc = f(n - 1)
                       Else
                       sc = sc + f(n - 1)
                       End If
                    Next n
            
                    Set line = ThisDrawing.ModelSpace.AddLine(g, u)
                        x(0) = varPick(0) + (sm) + (sc)
                        x(1) = varPick(1)
                    
                    y(0) = x(0)
                    y(1) = x(1) + 186.5455
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(x, y)
                    line.Update
                    
                    Dim adi(0 To 2) As Double
                    adi(0) = x(0)
                    adi(1) = x(1) - 25 ' notice here 25 is cover lenght...
'                    Set line = ThisDrawing.ModelSpace.AddLine(adi, x)
'                    line.Update

                    nio = 3
                    For nis = 1 To nio
                        Dim aki(0 To 2) As Double
                        aki(0) = adi(0)
                        aki(1) = adi(1) - 200 + 25 + 25
                        Set line = ThisDrawing.ModelSpace.AddLine(aki, adi)
                        line.Update
                            
                        adi(0) = adi(0) + 120
                        adi(1) = adi(1)
                        
                    Next nis
                        
                    
                    
                    
            ' ___________________________________ top top top top .....
                    ' column width is = = = f(n-1)
                    Dim eee(0 To 2) As Double
                    eee(0) = i(0) + (((f(n - 2)) / 2) - (73.418 / 2))
                    eee(1) = i(1)
                    
                    jis(0) = x(0)
                    jis(1) = x(1) - (200 / 2)
'                     Set line = ThisDrawing.ModelSpace.AddLine(jis, x)
                     ' line.update
                                                            
                    Set line = ThisDrawing.ModelSpace.AddLine(eee, i)
                    line.Update
                    Dim aaa(0 To 2) As Double
                        aaa(0) = eee(0) + 24.1052
                        aaa(1) = eee(1) + 70.9835
                        
                   Set line = ThisDrawing.ModelSpace.AddLine(aaa, eee)
                line.Update
                
                 Dim fff(0 To 2) As Double
                    fff(0) = y(0) - ((f(n - 2) / 2) - (73.418 / 2))
                    fff(1) = y(1)
                    
                   Set line = ThisDrawing.ModelSpace.AddLine(fff, y)
                line.Update
                    
                 Dim bbb(0 To 2) As Double
                        bbb(0) = fff(0) - 20.3504
                        bbb(1) = fff(1) - 38.8248
                
                        Set line = ThisDrawing.ModelSpace.AddLine(fff, bbb)
                        line.Update
                        Set line = ThisDrawing.ModelSpace.AddLine(bbb, aaa)
                        line.Update
              Dim jjj(0 To 2) As Double
                    jjj(0) = i(0) - 60
                    jjj(1) = i(1)
               Dim yyy(0 To 2) As Double
                    yyy(0) = y(0) + 60
                    yyy(1) = y(1)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(jjj, i)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(yyy, y)
                    line.Update

                    tca(0) = i(0) + (f(n - 2) / 2)
                    tca(1) = i(1)

                    ' b straight line up..coodinate..

                    'to find b coordinate
                    tcb(0) = tca(0)
                    tcb(1) = tca(1) + 335

                    Set line = ThisDrawing.ModelSpace.AddLine(tcb, tca)

                    ' c center line from  radius of circle find coodinate

                    tcc(0) = tcb(0)
                    tcc(1) = tcb(1) + 194

                    ' ... draw circle.. with the above coordinate....tcc(0) and tcc(1)

                    ' ????????????????????????????
                    
                    Set objEnt = ThisDrawing.ModelSpace.AddCircle(tcc, 194)
                    line.Update

                    ' d coodinate frm b coodinate... add dia of circle
                    tcd(0) = tcc(0)
                    tcd(1) = tcc(1) + 194

                    ' e coodinate frm d coodinate... add dia of dist_ed
                    tce(0) = tcd(0)
                    tce(1) = tcd(1) + 147

                    Set line = ThisDrawing.ModelSpace.AddLine(tcd, tce)


                    ' f coodinate frm c centerCircle..
                    tcf(0) = tcc(0) - 194
                    tcf(1) = tcc(1)

                    ' h coordinate frm f coordinate add dist_fh
                    tch(0) = tcf(0) - 160
                    tch(1) = tcf(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(tch, tcf)
                    ' g coodinate frm c centerCircle ..
                    tcg(0) = tcc(0) + 194
                    tcg(1) = tcc(1)

                    ' i coodinate frm g coordinate add dist_gi

                    tci(0) = tcg(0) + 160
                    tci(1) = tcg(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(tci, tcg)
                    line.Update
                    'add text in the cirlcle using the center coodinate...
                    ' text that upddate itself again again... make use of the loop..
                    ' ????????????????????????????
                    
                    height = 200

                    tcj(0) = tcc(0) - 69.6153
                    tcj(1) = tcc(1) + 90.4984

                    Set textObj = ThisDrawing.ModelSpace.AddMText(tcj, 200, sn)
                    textObj.height = 200
                    ' Set mtextObj = ThisDrawing.ModelSpace.AddMText(insertPoint, width, textString)
                    line.Update
                    'some how check to see if the text is inside the circle....
                    
                '4. get e(0) ...coodinate in similar fashion as the above..  f(n-2) center
                    cc(0) = e(0) + (f(n - 2) / 2) ' here add column width c1
                    cc(1) = e(1)
                    
                    ci(0) = i(0) + (f(n - 2) / 2) ' here add column width c1
                    ci(1) = i(1)
                
                '5.with dist s= 305 find coodinate... call it ...di(0)... dio(1)...
                   
                   Dim dii(0 To 2) As Double
                    dii(0) = cc(0) ' here add column width c1
                    dii(1) = cc(1) - 305
                    
                '3. add line... with dist = 500+305 fm center of column...coodinate...
                    die(0) = dii(0) ' here add column width c1
                    die(1) = dii(1) - 500
                
                    Set line = ThisDrawing.ModelSpace.AddLine(die, ci)
                    ' 9.9736
                    line.Update
                                ' Dim mSp As AcadModelSpace
                                            
                                
                                Dim sDim As AcadDimRotated
                                                                    
                                    If sn = 2 Then

                                        Dim pts1(0 To 2) As Double
                                        Dim pts2(0 To 2) As Double
                                        Dim loc(0 To 2) As Double
                                        

                                        pts1(0) = dis(0)
                                        pts1(1) = dis(1)

                                        pts2(0) = dii(0)
                                        pts2(1) = dii(1)

                                        loc(0) = p(0) / 2
                                        loc(1) = pts1(1)

                                        rotAngle = 0
                                        rotAngle = rotAngle * 3.141592 / 180#       ' covert to Radians
                                                                        
                                        'Add dimension
                                        Set sDim = ThisDrawing.ModelSpace.AddDimRotated(pts1(), pts2(), loc(), rotAngle)

                                        'Set dimension properties
                                        sDim.color = acByLayer

                                        'sDim.ExtensionLineExtend = 0

                                        sDim.LinetypeScale = 100

                                        sDim.Arrowhead1Type = acArrowArchTick
                                        sDim.Arrowhead2Type = acArrowArchTick
                                        '        sDim.arrowsize
                                        sDim.ArrowheadSize = 100
                                        sDim.TextColor = RGB(255, 127, 0)
                                        ' sDim.TextColor = RGB(255, 127, 0)
                                        sDim.TextHeight = 200
                                        ' sDim.TextHeight = 220
                                        sDim.UnitsFormat = acDimLDecimal

                                        sDim.ExtLine1Suppress = True
                                        sDim.ExtLine2Suppress = True

                                        sDim.PrimaryUnitsPrecision = acDimPrecisionOne
                                        sDim.TextGap = 30
                                        ' sDim.TextGap = 3
                                        sDim.LinearScaleFactor = 1
                                        ' sDim.LinearScaleFactor = 1
                                        sDim.ExtensionLineOffset = 0
                                        ' sDim.ExtensionLineOffset = 1000
                                        ThisDrawing.ActiveTextStyle.height = 180
                                        sDim.VerticalTextPosition = acAbove
                                        ' sDim.VerticalTextPosition = acAbove

                                        sDim.PrimaryUnitsPrecision = acDimPrecisionZero
                                        'Create a new dimension style
                                        Set dimstyle = ThisDrawing.DimStyles.Add("D100")

                                        'Create a new dimension style
                                        'Set dimstyle = ThisDrawing.DimStyles.Add("jjkj")
                                        sDim.Update
                                    
                                    Else

            sn = sn - 2
                
                If sn = 1 Then
                    ' need some sort of if here like the above one...
                        ' if sn=1... proceed... with this step..
                        ' if sn=>1... proceed with for loop...as in the former... one...
                    n = sn + 1
                        ' then find pts1.. ie..die(0)
                            x(0) = varPick(0)
                            x(1) = varPick(1)

                            y(0) = x(0)
                            y(1) = x(1) + 186.5455

                            Set line = ThisDrawing.ModelSpace.AddLine(x, y)

                            z(0) = y(0) + p(sn - 1)
                            z(1) = x(1)

                            Set line = ThisDrawing.ModelSpace.AddLine(x, z)

                            i(0) = z(0)
                            i(1) = z(1) + 186.5455

                            Set line = ThisDrawing.ModelSpace.AddLine(z, i)

                            
                            ' here is where ... beam width is  added..b1 say

                            b(0) = y(0)
                            b(1) = x(1) - 200 '.....the length
                ' ________________________ 200 here is the beam_width....

                            c(0) = y(0)
                            c(1) = b(1) - 186.5455

                        '    _____186.545.. is the length b/n c-b...
                            
                            Set line = ThisDrawing.ModelSpace.AddLine(c, b)
                            
                            
                            d(0) = b(0) + p(sn - 1) ' here is where l(1) goes
                            d(1) = b(1)
                            Set line = ThisDrawing.ModelSpace.AddLine(b, d)
                            line.Update
                            
                            e(0) = d(0)
                            e(1) = d(1) - 186.5455
                 
                '4. get e(0) ...coodinate in similar fashion as the above..  f(n-2) center
                                                
                            cc(0) = e(0) + (f(n - 1) / 2) ' here add column width c1
                            cc(1) = e(1)
                            
                            ci(0) = i(0) + (f(n - 1) / 2) ' here add column width c1
                            ci(1) = i(1)
                        
                '5.with dist s= 305 find coodinate... call it ...di(0)... dio(1)...
                           
                           Dim diii(0 To 2) As Double
                            diii(0) = cc(0) ' here add column width c1
                            diii(1) = cc(1) - 305
                                ' pts1(0)
                
                '3. add line... with dist = 500+305 fm center of column...coodinate...
                            die(0) = diii(0) ' here add column width c1
                            die(1) = diii(1) - 500
                        
                            Set line = ThisDrawing.ModelSpace.AddLine(die, ci)
                                line.Update
                                
            Else
'                 n = sn + 1
                For m = 2 To sn
                   If m = 2 Then
                       sm = p(m - 2)
                   Else
                   sm = sm + p(m - 2)
                   End If
                Next m
                
                For n = 2 To sn
                   If n = 2 Then
                       sc = f(n - 1)
                   Else
                   sc = sc + f(n - 1)
                   End If
                Next n



                Set line = ThisDrawing.ModelSpace.AddLine(g, u)
                    x(0) = varPick(0) + (sm) + (sc)
                    x(1) = varPick(1)
                
                y(0) = x(0)
                y(1) = x(1) + 186.5455
                
                Set line = ThisDrawing.ModelSpace.AddLine(x, y)
                line.Update
                                    
'                 n = sn + 1
                
                   line.Update
                    z(0) = y(0) + p(sn - 1)
                    z(1) = x(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(x, z)
                    line.Update

                    
                    
                    i(0) = z(0)
                    i(1) = z(1) + 186.5455

                    Set line = ThisDrawing.ModelSpace.AddLine(z, i)
                    line.Update
                    b(0) = y(0)
                    b(1) = x(1) - 200 '.....the length..


                    c(0) = y(0)
                    c(1) = b(1) - 186.5455
                    
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(c, b)
                    
                    line.Update


                    d(0) = b(0) + p(sn - 1) ' here is where l(1) goes
                    d(1) = b(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(b, d)
                   line.Update
                    e(0) = d(0)
                    e(1) = d(1) - 186.5455

                    Set line = ThisDrawing.ModelSpace.AddLine(d, e)
                       line.Update
                     n = sn + 2
                    cc(0) = e(0) + (f(n - 2) / 2) ' here add column width c1
                    cc(1) = e(1)
                    
                    ci(0) = i(0) + (f(n - 2) / 2) ' here add column width c1
                    ci(1) = i(1)
                
                    '5.with dist s= 305 find coodinate... call it ...di(0)... dio(1)...
                               
                               
                                diii(0) = cc(0) ' here add column width c1
                                diii(1) = cc(1) - 305
                                
                    
                    '3. add line... with dist = 500+305 fm center of column...coodinate...
                                die(0) = diii(0) ' here add column width c1
                                die(1) = diii(1) - 500
                            
                                Set line = ThisDrawing.ModelSpace.AddLine(die, ci)
                                ' 9.9736
                      line.Update
                                                                         
            End If
              
              sn = sn + 2
            
                    pts1(0) = diii(0)
                    pts1(1) = diii(1)

                    pts2(0) = dii(0)
                    pts2(1) = dii(1)

                    loc(0) = p(0) / 2
                    loc(1) = pts1(1)

                    rotAngle = 0
                    rotAngle = rotAngle * 3.141592 / 180#       ' covert to Radians
                                                    
                    'Add dimension
                    Set sDim = ThisDrawing.ModelSpace.AddDimRotated(pts1(), pts2(), loc(), rotAngle)

                    'Set dimension properties
                    sDim.color = acByLayer

                    ' sDim.ExtensionLineExtend = 0

                    sDim.LinetypeScale = 100

                    sDim.Arrowhead1Type = acArrowArchTick
                    sDim.Arrowhead2Type = acArrowArchTick
                    '        sDim.arrowsize
                    sDim.ArrowheadSize = 100
                    sDim.TextColor = RGB(255, 127, 0)
                    ' sDim.TextColor = RGB(255, 127, 0)
                    sDim.TextHeight = 200
                    ' sDim.TextHeight = 220
                    sDim.UnitsFormat = acDimLDecimal

                    sDim.ExtLine1Suppress = True
                    sDim.ExtLine2Suppress = True

                    sDim.PrimaryUnitsPrecision = acDimPrecisionOne
                    sDim.TextGap = 30
                    ' sDim.TextGap = 3
                    sDim.LinearScaleFactor = 1
                    ' sDim.LinearScaleFactor = 1
                    sDim.ExtensionLineOffset = 0
                    ' sDim.ExtensionLineOffset = 1000
                    ThisDrawing.ActiveTextStyle.height = 180
                    sDim.VerticalTextPosition = acAbove
                    ' sDim.VerticalTextPosition = acAbove

                    sDim.PrimaryUnitsPrecision = acDimPrecisionZero
                    'Create a new dimension style
                    Set dimstyle = ThisDrawing.DimStyles.Add("D100")

                    'Create a new dimension style
                    ' Set dimstyle = ThisDrawing.DimStyles.Add("jjkj")
                    sDim.Update


              sn = sn - 1
            
                        For m = 2 To sn
                           If m = 2 Then
                               sm = p(m - 2)
                           Else
                           sm = sm + p(m - 2)
                           End If
                        Next m
                        
                        For n = 2 To sn
                           If n = 2 Then
                               sc = f(n - 1)
                           Else
                           sc = sc + f(n - 1)
                           End If
                        Next n
                
                        Set line = ThisDrawing.ModelSpace.AddLine(g, u)
                            x(0) = varPick(0) + (sm) + (sc)
                            x(1) = varPick(1)
                        
                        y(0) = x(0)
                        y(1) = x(1) + 186.5455
                    
                    
                    b(0) = y(0)
                    b(1) = x(1) - 200 '.....the length..
    
                    d(0) = b(0) + p(sn - 1) ' here is where l(1) goes
                    d(1) = b(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(b, d)
                   line.Update
                    e(0) = d(0)
                    e(1) = d(1) - 186.5455
                    Set line = ThisDrawing.ModelSpace.AddLine(d, e)
                        line.Update
             
             
             sn = sn + 1
                        For m = 2 To sn
                           If m = 2 Then
                               sm = p(m - 2)
                           Else
                           sm = sm + p(m - 2)
                           End If
                        Next m
                        
                        For n = 2 To sn
                           If n = 2 Then
                               sc = f(n - 1)
                           Else
                           sc = sc + f(n - 1)
                           End If
                        Next n
                
                        Set line = ThisDrawing.ModelSpace.AddLine(g, u)
                            x(0) = varPick(0) + (sm) + (sc)
                            x(1) = varPick(1)
                        
                        y(0) = x(0)
                        y(1) = x(1) + 186.5455

                        Set line = ThisDrawing.ModelSpace.AddLine(x, y)
                        line.Update
    
                End If
                        
                    
                    line.Update
                    z(0) = y(0) + p(sn - 1)
                    z(1) = x(1)
                    
                    Dim dzo(0 To 2) As Double
                    dzo(0) = z(0)
                    dzo(1) = z(1) - 25
'                    Set line = ThisDrawing.ModelSpace.AddLine(dzi, z)
'                    line.Update
                                    
                    ' loop for sheear rein
                    nio = 3
                    For nis = 1 To nio
                        Dim sho(0 To 2) As Double
                        sho(0) = dzo(0)
                        sho(1) = dzo(1) - 200 + 25 + 25
                        Set line = ThisDrawing.ModelSpace.AddLine(sho, dzo)
                        line.Update
                            
                        dzo(0) = dzo(0) - 120
                        dzo(1) = dzo(1)
                            
                    Next nis
                                    
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(x, z)
                    line.Update
                    i(0) = z(0)
                    i(1) = z(1) + 186.5455

                    jii(0) = z(0)
                    jii(1) = z(1) - (200 / 2)
                     ' Set line = ThisDrawing.ModelSpace.AddLine(jii, z)
'                     line.Update

                  ' add new dimension properties here for jis() & jii()..
                    
                    ' here add dim variables
                    
                    ptc1(0) = jis(0)
                    ptc1(1) = jis(1)
'                      Set line = ThisDrawing.ModelSpace.AddLine(jii, z)
'                     line.Update

                    ptc2(0) = jii(0)
                    ptc2(1) = jii(1)

                    lod(0) = jis(0) / 2
                    lod(1) = ptc1(1)

                    rotAngle = 0
                    rotAngle = rotAngle * 3.141592 / 180#       ' covert to Radians
                                                                                    
                    'Add dimension
                    Set aDim = ThisDrawing.ModelSpace.AddDimRotated(ptc1(), ptc2(), lod(), rotAngle)

                    'Set dimension properties
                    aDim.color = acByLayer

                    'aDim.ExtensionLineExtend = 0

                    aDim.LinetypeScale = 100

                    aDim.Arrowhead1Type = acclosedfilled
                    aDim.Arrowhead2Type = acclosedfilled
                    '        aDim.arrowsize
                    aDim.ArrowheadSize = 100
                    aDim.TextColor = RGB(255, 127, 0)
                    ' aDim.TextColor = RGB(255, 127, 0)
                    'notice here 200 = beam width
                
                    aDim.TextHeight = 85
                    ' aDim.TextHeight = 220
                    aDim.UnitsFormat = acDimLDecimal

                    aDim.ExtLine1Suppress = True
                    aDim.ExtLine2Suppress = True

                    aDim.PrimaryUnitsPrecision = acDimPrecisionOne
                    aDim.TextGap = 30
                    ' aDim.TextGap = 3
                    aDim.LinearScaleFactor = 1
                    ' aDim.LinearScaleFactor = 1
                    aDim.ExtensionLineOffset = 0
                    ' aDim.ExtensionLineOffset = 1000

                    aDim.VerticalTextPosition = acCentered
                    ' aDim.VerticalTextPosition = acAbove
                    ThisDrawing.ActiveTextStyle.height = 85
                    aDim.PrimaryUnitsPrecision = acDimPrecisionZero
                    'Create a new dimension style
                    
                    aDim.TextOverride = "{\fUtsaah|b0|i0|c0|p34;8@c/c180}"
                    'aDim.TextStyle = sdf

                    Set dimstyle = ThisDrawing.DimStyles.Add("D100")

                    'Create a new dimension style
                    'Set dimstyle = ThisDrawing.DimStyles.Add("jjkj")


                    aDim.Update
                     
                    Set line = ThisDrawing.ModelSpace.AddLine(z, i)
                    line.Update
' ___________________________________  top top top top.....
                    
                    ' here is where ... beam width is  added..b1 say
                    b(0) = y(0)
                    b(1) = x(1) - 200 '.....the length..
                    
                '????????????//////    _______ 200 here is the beam_width....
                    c(0) = y(0)
                    c(1) = b(1) - 186.5455
                      
                    Set line = ThisDrawing.ModelSpace.AddLine(c, b)
                    line.Update
' ___________________________________  bottom bottom bottom bottom.....

                    ' ....that starts with ....c... same as ...k....
                     Dim cce(0 To 2) As Double
                    cce(0) = c(0) - ((f(n - 2) / 2) - (73.418 / 2))
                    cce(1) = c(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(cce, c)
                    Dim chh(0 To 2) As Double
                    
                        chh(0) = cce(0) - 20.3504
                        chh(1) = cce(1) - 38.8248
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(chh, cce)
                line.Update
                    
                    ' ... that starts with ....e... same as ...k....
                    
                    Dim eec(0 To 2) As Double
                    eec(0) = e(0) + ((f(n - 2) / 2) - (73.418 / 2))
                    eec(1) = e(1)
                    

                    Set line = ThisDrawing.ModelSpace.AddLine(eec, e)
                line.Update
                       
                   Dim eeg(0 To 2) As Double
                        eeg(0) = eec(0) + 24.1052
                        eeg(1) = eec(1) + 70.9835
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(eeg, eec)
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(eeg, chh)
    
                        line.Update
    
            ' need to add constants around here for the bottom
                    
                    jj(0) = e(0) - 60
                    jj(1) = e(1)
                    
                    yy(0) = c(0) + 60
                    yy(1) = c(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(jj, e)
                    Set line = ThisDrawing.ModelSpace.AddLine(yy, c)
                    line.Update
    
                    d(0) = b(0) + p(sn - 1) ' here is where l(1) goes
                    d(1) = b(1)
                    Set line = ThisDrawing.ModelSpace.AddLine(b, d)
                   line.Update
                    e(0) = d(0)
                    e(1) = d(1) - 186.5455

                    Set line = ThisDrawing.ModelSpace.AddLine(d, e)
                    line.Update
                        
                        
                    

                
                
        ' ___________________________________  bottom bottom bottom bottom.....
    End If
    
Next sn

    g(0) = e(0) + f(n - 1) ' here add column width c of the last one....
    g(1) = i(1)
    u(0) = e(0) + f(n - 1) 'here add c1
    u(1) = e(1)  '  jtotal = 2*186.5455 + b1 here add c1
    Set line = ThisDrawing.ModelSpace.AddLine(g, u)
                    ' column width is = = = f(n-1)
                    
                            Dim aoi(0 To 2) As Double
                            aoi(0) = z(0)
                            aoi(1) = z(1) - 25 ' notice here 25 is cover lenght...
                            
'                            Set line = ThisDrawing.ModelSpace.AddLine(aoi, z)
'                            line.Update
                                
                            Dim aoj(0 To 2) As Double
                            aoj(0) = z(0) + f(n - 1) - 25
                            aoj(1) = aoi(1)
                            Set line = ThisDrawing.ModelSpace.AddLine(aoj, aoi)
                            line.Update
                            
                            Dim aox(0 To 2) As Double
                            aox(0) = z(0)
                            aox(1) = z(1) - (200) + 25 'notice here 200 is beam width
'                            Set line = ThisDrawing.ModelSpace.AddLine(aox, aoi)
'                            line.Update
                            
                            Dim aoy(0 To 2) As Double
                            aoy(0) = z(0) + f(n - 1) - 25
                            aoy(1) = aox(1)
                            Set line = ThisDrawing.ModelSpace.AddLine(aoy, aoj)
                            line.Update
                                
'                             here add joining lines
                            Set line = ThisDrawing.ModelSpace.AddLine(asj, aoi)
                            line.Update
                            
                            Set line = ThisDrawing.ModelSpace.AddLine(asy, aoy)
                            
                            line.Update
                                
                                
                                
                    
                    Dim eue(0 To 2) As Double
                    eue(0) = i(0) + (((f(n - 1)) / 2) - (73.418 / 2))
                    eue(1) = i(1)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(eue, i)
                    Dim jio(0 To 2) As Double
                    jio(0) = i(0)
                    jio(1) = i(1) - 186.5455 - (200 / 2)
                    ' Set line = ThisDrawing.ModelSpace.AddLine(jio, i)
'                    line.Update
                     Dim aua(0 To 2) As Double
                        aua(0) = eue(0) + 24.1052
                        aua(1) = eue(1) + 70.9835
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(aua, eue)
        
                     Dim fuf(0 To 2) As Double
                    fuf(0) = g(0) - ((f(n - 1) / 2) - (73.418 / 2))
                    fuf(1) = g(1)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(fuf, g)

                    
                    Dim bub(0 To 2) As Double
                        bub(0) = fuf(0) - 20.3504
                        bub(1) = fuf(1) - 38.8248
                
                        Set line = ThisDrawing.ModelSpace.AddLine(fuf, bub)
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(bub, aua)
                        
                   Dim juj(0 To 2) As Double
                    juj(0) = i(0) - 60
                    juj(1) = i(1)
                   
                    Dim yuy(0 To 2) As Double
                    yuy(0) = g(0) + 60
                    yuy(1) = g(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(juj, i)
                    
                    Set line = ThisDrawing.ModelSpace.AddLine(yuy, g)
                    
                    tca(0) = i(0) + (f(n - 1) / 2)
                    tca(1) = i(1)

                    ' b straight line up..coodinate..

                    'to find b coordinate
                    tcb(0) = tca(0)
                    tcb(1) = tca(1) + 335

                    Set line = ThisDrawing.ModelSpace.AddLine(tcb, tca)

                    ' c center line from  radius of circle find coodinate
                    tcc(0) = tcb(0)
                    tcc(1) = tcb(1) + 194

                    ' ... draw circle.. with the above coordinate....tcc(0) and tcc(1)

                    ' ????????????????????????????
                    
                    Set objEnt = ThisDrawing.ModelSpace.AddCircle(tcc, 194)


                    ' d coodinate frm b coodinate... add dia of circle
                    tcd(0) = tcc(0)
                    tcd(1) = tcc(1) + 194

                    ' e coodinate frm d coodinate... add dia of dist_ed
                    tce(0) = tcd(0)
                    tce(1) = tcd(1) + 147

                    Set line = ThisDrawing.ModelSpace.AddLine(tcd, tce)


                    ' f coodinate frm c centerCircle..
                    tcf(0) = tcc(0) - 194
                    tcf(1) = tcc(1)

                    ' h coordinate frm f coordinate add dist_fh
                    tch(0) = tcf(0) - 160
                    tch(1) = tcf(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(tch, tcf)

                    ' g coodinate frm c centerCircle ..
                    tcg(0) = tcc(0) + 194
                    tcg(1) = tcc(1)


                    ' i coodinate frm g coordinate add dist_gi

                    tci(0) = tcg(0) + 160
                    tci(1) = tcg(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(tci, tcg)

                    'add text in the cirlcle using the center coodinate...
                    ' text that upddate itself again again... make use of the loop..
                    ' ????????????????????????????

                    height = 200

                    tcj(0) = tcc(0) - 69.6153
                    tcj(1) = tcc(1) + 90.4984


                    Set textObj = ThisDrawing.ModelSpace.AddMText(tcj, 200, sn)
                    textObj.height = 200
                    ' Set mtextObj = ThisDrawing.ModelSpace.AddMText(insertPoint, width, textString)

                    'some how check to see if the text is inside the circle....

                    ' 9.9736

                    line.Update

                     Dim cue(0 To 2) As Double
                    cue(0) = u(0) - ((f(n - 1) / 2) - (73.418 / 2))
                    cue(1) = u(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(cue, u)
                    
                     Dim cuh(0 To 2) As Double
                        cuh(0) = cue(0) - 20.3504
                        cuh(1) = cue(1) - 38.8248
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(cuh, cue)
                
                    
                    ' ... that starts with ....e... same as ...k....
                   Dim euc(0 To 2) As Double
                    euc(0) = e(0) + ((f(n - 1) / 2) - (73.418 / 2))
                    euc(1) = e(1)
                    

                    Set line = ThisDrawing.ModelSpace.AddLine(euc, e)
                    Dim eug(0 To 2) As Double
                        eug(0) = euc(0) + 24.1052
                        eug(1) = euc(1) + 70.9835
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(eug, euc)
                        
                        Set line = ThisDrawing.ModelSpace.AddLine(eug, cuh)
                    
                      Dim jgj(0 To 2) As Double
                    jgj(0) = e(0) - 60
                    jgj(1) = e(1)
                    
                 Dim ygy(0 To 2) As Double
                    ygy(0) = u(0) + 60
                    ygy(1) = u(1)

                    Set line = ThisDrawing.ModelSpace.AddLine(jgj, e)
                    Set line = ThisDrawing.ModelSpace.AddLine(ygy, u)
                    

            '4.get e(0) ...coodinate in similar fashion as the above..  f(n-1) center
                   
                    cc(0) = e(0) + (f(n - 1) / 2) ' here add column width c1
                    cc(1) = e(1)
                    
                    ci(0) = i(0) + (f(n - 1) / 2) ' here add column width c1
                    ci(1) = i(1)
                                        
        '5.with dist s= 305 find coodinate... call it ...di(0)... dio(1)...
                   
                    Dim dio(0 To 2) As Double
                    dio(0) = cc(0) ' here add column width c1
                    dio(1) = cc(1) - 305
                            
        '3. add line... with dist = 500+305 fm center of column...coodinate...
                    
                    die(0) = dio(0) ' here add column width c1
                    die(1) = dio(1) - 500
                
                    Set line = ThisDrawing.ModelSpace.AddLine(die, ci)
                    line.Update
  
                                        pts1(0) = dii(0)
                                        pts1(1) = dii(1)

                                        pts2(0) = dio(0)
                                        pts2(1) = dio(1)

                                        loc(0) = p(0) / 2
                                        loc(1) = pts1(1)

                                        rotAngle = 0
                                        rotAngle = rotAngle * 3.141592 / 180#       ' covert to Radians
                                                                        
                                        'Add dimension
                                        Set sDim = ThisDrawing.ModelSpace.AddDimRotated(pts1(), pts2(), loc(), rotAngle)

                                        'Set dimension properties
                                        sDim.color = acByLayer

                                        'sDim.ExtensionLineExtend = 0

                                        sDim.LinetypeScale = 100

                                        sDim.Arrowhead1Type = acArrowArchTick
                                        sDim.Arrowhead2Type = acArrowArchTick
                                        '        sDim.arrowsize
                                        sDim.ArrowheadSize = 100
                                        sDim.TextColor = RGB(255, 127, 0)
                                        ' sDim.TextColor = RGB(255, 127, 0)
                                        sDim.TextHeight = 200
                                        ' sDim.TextHeight = 220
                                        sDim.UnitsFormat = acDimLDecimal

                                        sDim.ExtLine1Suppress = True
                                        sDim.ExtLine2Suppress = True

                                        sDim.PrimaryUnitsPrecision = acDimPrecisionOne
                                        sDim.TextGap = 30
                                        ' sDim.TextGap = 3
                                        sDim.LinearScaleFactor = 1
                                        ' sDim.LinearScaleFactor = 1
                                        sDim.ExtensionLineOffset = 0
                                        ' sDim.ExtensionLineOffset = 1000
                                        ThisDrawing.ActiveTextStyle.height = 180
                                        sDim.VerticalTextPosition = acAbove
                                        ' sDim.VerticalTextPosition = acAbove

                                        sDim.PrimaryUnitsPrecision = acDimPrecisionZero
                                        'Create a new dimension style
                                        Set dimstyle = ThisDrawing.DimStyles.Add("D100")

                                        'Create a new dimension style
                                        '                                        Set dimstyle = ThisDrawing.DimStyles.Add("jjkj")
                                        sDim.Update
     
End With

' rebar dwg...

            Dim ra(0 To 2) As Double
                ra(0) = asj(0)
                ra(1) = asj(1) - 1695.1543
'            Set line = ThisDrawing.ModelSpace.AddLine(ra, asj)
'            line.Update

            ' 2.with ra(0).. go a distance of (bw-2*cover) down..rax(0)
                ' and add a curve that with degree of 45 with
                ' distance of.. with rai(0) var

            ' going distance down
                Dim disb(0 To 2) As Double
                    disb(0) = ra(0)
                    disb(1) = ra(1) + 350 'rbs=sec - (25+25)
                Set line = ThisDrawing.ModelSpace.AddLine(ra, disb)
                line.Update
            ' doing the curver
                Dim dicu(0 To 2) As Double
                    dicu(0) = disb(0) + 128.774
                    dicu(1) = ra(1) + 67.878
                Set line = ThisDrawing.ModelSpace.AddLine(ra, dicu)
                line.Update

                ' *span_1 + span_2...+ rbs < 12 kind of create a loop here...
                        ' such that the loop here would do is..cumulitative kind of thing
                        
                     o = 5
'                o = 5
                
                For pn = 1 To o
'                    g(pn) = f(pn)
'                    k(pn) = p(pn)
                    
                    sum = p(pn - 1) + p(pn) '+ constants
                    ' sum= t1 + t2 + constants
                        ' t1=t2
                         p(pn - 1) = p(pn)
                         p(pn) = sum

                    gum = f(pn - 1) + f(pn) '+ constants
                    ' sum= t1 + t2 + constants
                        ' t1=t2
                         f(pn - 1) = f(pn)
                         f(pn) = gum
                    
                    fnl = gum + sum - 25
                        ' 14675
                    If fnl < 12000 Then
                '            exit if
                        Else
                        
                        
'                             u need a select case here


                            pn = pn - 1
                            half = s(pn) / 2

                            rev = fnl + half - s(pn + 1) - s(pn) - h(pn + 1)
                                                        
                                Dim conkt(0 To 2) As Double
                                    conkt(0) = disb(0) + rev
                                    conkt(1) = disb(1)
                                Set line = ThisDrawing.ModelSpace.AddLine(conkt, disb)
                                line.Update
                                ' do the curve going a distace..
'                                        going a dista of half overlap length
'                                        overlap length and do the curve
                                
                                    Dim dstOv(0 To 2) As Double
                                        dstOv(0) = conkt(0) + 250 'overlapDist
                                        dstOv(1) = conkt(1)
                                    Set line = ThisDrawing.ModelSpace.AddLine(conkt, dstOv)
                                    line.Update
                                    
                                    'do the curve
                                    Dim concve(0 To 2) As Double
                                        concve(0) = dstOv(0) - 128.774
                                        concve(1) = dstOv(1) - 67.878
                                    Set line = ThisDrawing.ModelSpace.AddLine(dstOv, concve)
                                    line.Update


                                ' go a distance up of 148.2
                                Dim conki(0 To 2) As Double
                                    conki(0) = conkt(0)
                                    conki(1) = conkt(1) + 148.2
                                Set line = ThisDrawing.ModelSpace.AddLine(conkt, conki)
'                                line.Update
                                                                
                                ' go overlap distance to left...
                                Dim dsOv(0 To 2) As Double
                                    dsOv(0) = conki(0) - 250 'overlapDist
                                    dsOv(1) = conki(1)
                                Set line = ThisDrawing.ModelSpace.AddLine(dsOv, conki)
                                line.Update
                                                            
'                                do the curve
                                 Dim concvx(0 To 2) As Double
                                        concvx(0) = dsOv(0) + 128.774
                                        concvx(1) = dsOv(1) - 67.878
                                 Set line = ThisDrawing.ModelSpace.AddLine(dsOv, concvx)
                                 line.Update
                                
'                                start the next loop here... using the coordinate..of following
                                
                                
                                
'                                using conki(0) coordinate
                                
                                
                                
                                
                                
                                ' Loop for sum not new but use the last rev one...
                                
                            ' some how try making 12 to 24 and .. other
                            ' rev =
                            ' pn=4
                            ' go to
                            
                    End If
                                                
                Next pn


                
                    
                ' *then find the previous from the last span..
                ' *try doing the overlap.. here..dependent on dimeter of bar
                    
                    ' *......part_1........
                    ' *go at a distance of overlap/2.. and do the curve
                    ' *at the half of the journey find the coordinate().. do the curve
                    ' *go up at distance of 148.2 up and go a dist of overlap/2 back and coordinate()
                    ' *do the curve then
                    
                    ' *......part_1_End........
                    ' *from there   test again span_3 + span_4 + recentSpan/2..<12
                        ' how to know the last section part?
                            ' some how find a relation with the loop Safie .. has
                    ' *go a distance down of 148.2
                        ' *repeat part_1...
                        
                
'    Else
        ' proceed just like ordinary
        
        
        
'    End If
        

' 4.


' 5.













End Sub











