sub dim_beam()

Dim pts1(0 To 2) As Double
Dim pts2(0 To 2) As Double
Dim loc(0 To 2) As Double
Dim rotAngle As Double

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

'                                        sDim.ExtensionLineExtend = 0

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

sDim.VerticalTextPosition = acAbove
' sDim.VerticalTextPosition = acAbove

sDim.PrimaryUnitsPrecision = acDimPrecisionZero
'Create a new dimension style
Set dimstyle = ThisDrawing.DimStyles.Add("D100")

'Create a new dimension style
'Set dimstyle = ThisDrawing.DimStyles.Add("jjkj")


sDim.Update








End sub







































