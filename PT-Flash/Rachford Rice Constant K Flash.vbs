Sub RRCheck()

Dim ZiArray()
Dim TcArray()
Dim PcArray()
Dim OmegaArray()
Dim KiArray()
Dim RRArray()
Dim RRPrimeArray()
Dim Bv As Double
Dim TotalRR As Double
Dim TotalRRPrime As Double
Dim Kmax As Double
Dim KmaxPole As Double
Dim Kmin As Double
Dim KminPole As Double
Dim PhaseXiComposition()
Dim PhaseYiComposition()
Dim Threshold As Double
Dim error As Double
Dim iteration As Double
Dim Nc As Double
Dim i As Double
Dim j As Double
Dim T As Double
Dim P As Double
Dim R As Double

Threshold = Range("SAOnePhaseThreshold").Value
T = Range("SATemp").Value
P = Range("SAPress").Value
R = Range("SARConst").Value
Nc = Range("SANc").Value - 1

'Creating array for Tc
ReDim TcArray(Nc)
For i = 0 To Nc
Range("SATableTcHead").Select
TcArray(i) = ActiveCell.Offset(i + 1, 0).Value
Next i

'Creating array for Pc
ReDim PcArray(Nc)
For i = 0 To Nc
Range("SATablePcHead").Select
PcArray(i) = ActiveCell.Offset(i + 1, 0).Value
Next i

'Creating array for Zi
ReDim ZiArray(Nc)
For i = 0 To Nc
Range("SATableZiHead").Select
ZiArray(i) = ActiveCell.Offset(i + 1, 0).Value
Next i

'Creating array for omega
ReDim OmegaArray(Nc)
For i = 0 To Nc
Range("SATableOmegaHead").Select
OmegaArray(i) = ActiveCell.Offset(i + 1, 0).Value
Next i

ReDim KiArray(Nc)
For i = 0 To Nc
KiArray(i) = (PcArray(i) / P) * (Exp(5.373 * (1 + OmegaArray(i)) * (1 - (TcArray(i) / T))))
Next i


    '*************************************************************************************************************************'
    '**************************************This is step 2 of algorithm of 2 phase PT Flash************************************
    '*************************************************************************************************************************'

Bv = 0.99
TotalRR = 1

'Estimating Kmax/Kmin Pole to handle under/over relaxation of Bv
Kmax = Application.WorksheetFunction.Max(KiArray)
KmaxPole = 1 / (1 - Kmax)
Kmin = Application.WorksheetFunction.Min(KiArray)
KminPole = 1 / (1 - Kmin)

'estimating RR function and derivative for each Ki and Bv
ReDim RRArray(Nc)
ReDim RRPrimeArray(Nc)

Do While Abs(TotalRR) > Threshold
    
    For i = 0 To Nc
        RRArray(i) = ((1 - KiArray(i)) * ZiArray(i)) / (1 - ((1 - KiArray(i)) * Bv))
        RRPrimeArray(i) = (((KiArray(i) - 1) ^ 2) * ZiArray(i)) / ((((KiArray(i) - 1) * Bv) + 1) ^ 2)
    Next i

    'computing total RR function
    TotalRR = 0
    TotalRRPrime = 0
    For i = 0 To Nc
        TotalRR = TotalRR + RRArray(i)
        TotalRRPrime = TotalRRPrime + RRPrimeArray(i)
    Next i

    'Updating Bv based on newton iteration
    Bv = Bv - (TotalRR / TotalRRPrime)

    'Checking for under/over relaxation of Bv
    If Bv < KmaxPole Then
        Bv = 0.5 * (Bv + KmaxPole)
    ElseIf Bv > KminPole Then
        Bv = 0.5 * (Bv + KminPole)
    End If

Loop

' estimating phase xi & yi composition from estimated Bv
ReDim PhaseXiComposition(Nc)
ReDim PhaseYiComposition(Nc)

For i = 0 To Nc
PhaseXiComposition(i) = ZiArray(i) / (1 + (Bv * (KiArray(i) - 1)))
PhaseYiComposition(i) = KiArray(i) * PhaseXiComposition(i)
Next i

For i = 0 To Nc
    Range("SATableXiHead").Select
    ActiveCell.Offset(i + 1, 0).Value = PhaseXiComposition(i)
Next i

For i = 0 To Nc
    Range("SATableYiHead").Select
    ActiveCell.Offset(i + 1, 0).Value = PhaseYiComposition(i)
Next i

'Range("P22").Value = Bv
End Sub


