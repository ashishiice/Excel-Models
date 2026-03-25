' ============================================================================
' Duration_Model.bas — VBA Macro Module for Duration_Calculator.xlsx
' Author: Bolt 🦞 for Ashish Prakash
' Purpose: Monte Carlo yield simulation, What-If scenarios, one-click reporting
' Import: Alt+F11 → File → Import File → select this .bas file
' Assign : Create a button on "What-If" sheet → Assign Macro → RunWhatIf
' ============================================================================

Option Explicit

' ──────────────────────────────────────────────
' Main entry point — run this from the button
' ──────────────────────────────────────────────
Sub RunWhatIf()
    Dim wsInputs As Worksheet
    Dim wsWhatIf As Worksheet
    Dim wsResults As Worksheet

    Set wsInputs  = ThisWorkbook.Sheets("Inputs")
    Set wsWhatIf  = ThisWorkbook.Sheets("What-If")
    Set wsResults = ThisWorkbook.Sheets("Results")

    Dim principal As Double
    Dim couponRate As Double
    Dim ytm As Double
    Dim tenor As Double
    Dim freq As Double

    principal  = wsInputs.Range("C6").Value
    couponRate = wsInputs.Range("C7").Value
    ytm        = wsInputs.Range("C9").Value
    tenor      = wsInputs.Range("C10").Value
    freq       = wsInputs.Range("C8").Value

    ' ── 1. Sensitivity table update ──────────────
    Dim yieldScenarios(24) As Double
    Dim tenorScenarios(14) As Double
    Dim i As Integer, j As Integer

    For i = 0 To 24
        yieldScenarios(i) = 0.05 + i * 0.0025
    Next i
    For j = 0 To 14
        tenorScenarios(j) = j + 1
    Next j

    ' Update What-If sheet with live modified duration values
    ' (Formulas already in cells — just recalculate by forcing calc)
    wsWhatIf.Calculate

    ' ── 2. Monte Carlo Simulation ─────────────────
    Dim nSims As Long
    nSims = 10000

    Dim meanYield As Double, volYield As Double
    meanYield = ytm
    volYield  = 0.01   ' 100bp annual vol

    Dim prices() As Double
    ReDim prices(1 To nSims)

    Dim r As Long, p As Long
    Dim z As Double, yieldPath As Double
    Dim cf As Double, pv As Double
    Dim nPeriods As Long
    nPeriods = Int(tenor * freq)

    Dim startTime As Double, endTime As Double
    startTime = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim batch As Long
    batch = 1000

    For r = 1 To nSims
        ' Generate yield path using GBM approximation
        z = Application.WorksheetFunction.NormSInv(Rnd())
        yieldPath = meanYield + volYield * z
        yieldPath = Application.WorksheetFunction.Max(yieldPath, 0.0001)

        ' Price the bond at this simulated yield
        pv = 0
        For p = 1 To nPeriods
            cf = principal * couponRate / freq
            If p = nPeriods Then cf = cf + principal
            pv = pv + cf / (1 + yieldPath / freq) ^ p
        Next p
        prices(r) = pv

        ' Progress update every batch
        If r Mod batch = 0 Then
            Application.StatusBar = "Simulation: " & r & " / " & nSims & " done..."
        End If
    Next r

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    endTime = Timer

    ' ── 3. Statistics output ─────────────────────
    Dim meanPrice As Double, stdPrice As Double
    Dim minPrice As Double, maxPrice As Double
    Dim sortedPrices() As Double
    ReDim sortedPrices(1 To nSims)
    For r = 1 To nSims
        sortedPrices(r) = prices(r)
    Next r

    ' Simple sort using bubble sort (fast enough for 10k)
    Dim temp As Double
    For i = 1 To nSims - 1
        For j = i + 1 To nSims
            If sortedPrices(i) > sortedPrices(j) Then
                temp = sortedPrices(i)
                sortedPrices(i) = sortedPrices(j)
                sortedPrices(j) = temp
            End If
        Next j
    Next i

    meanPrice = Application.WorksheetFunction.Average(prices)
    stdPrice  = Application.WorksheetFunction.StDev(prices)
    minPrice  = sortedPrices(1)
    maxPrice  = sortedPrices(nSims)

    Dim var95 As Double, cvar99 As Double
    var95  = sortedPrices(Int(0.05 * nSims))
    cvar99 = Application.WorksheetFunction.Average(Array(sortedPrices(Int(0.01 * nSims)), sortedPrices(1)))

    ' ── 4. Write output to What-If sheet ──────────
    Dim outRow As Long
    outRow = 35

    wsWhatIf.Cells(outRow, 2).Value = "MONTE CARLO SIMULATION RESULTS"
    wsWhatIf.Cells(outRow, 2).Font.Bold = True
    wsWhatIf.Cells(outRow, 2).Font.Size = 11
    wsWhatIf.Cells(outRow, 2).Interior.Color = RGB(31, 56, 100)
    wsWhatIf.Cells(outRow, 2).Font.Color = RGB(255, 255, 255)

    Dim statsData As Variant
    statsData = Array( _
        "Number of simulations", nSims, "", _
        "Mean price (₹ Cr)", Round(meanPrice, 2), "Average bond price across all scenarios", _
        "Std Dev price (₹ Cr)", Round(stdPrice, 2), "Volatility of bond price", _
        "Min price (₹ Cr)", Round(minPrice, 2), "Worst case price (1st percentile approx)", _
        "Max price (₹ Cr)", Round(maxPrice, 2), "Best case price (99th percentile approx)", _
        "VaR 95% (₹ Cr)", Round(meanPrice - var95, 2), "Value at Risk at 95% confidence — price drop from mean", _
        "CVaR 99% (₹ Cr)", Round(meanPrice - cvar99, 2), "Conditional VaR — avg loss beyond VaR", _
        "Yield vol used (%):", volYield * 100, "Annualised yield volatility assumption", _
        "Simulation time (sec):", Round(endTime - startTime, 2), "" _
    )

    Dim statRow As Long
    statRow = outRow + 2
    Dim k As Integer
    For k = 0 To UBound(statsData) Step 3
        wsWhatIf.Cells(statRow, 2).Value = statsData(k)
        wsWhatIf.Cells(statRow, 2).Font.Size = 10
        wsWhatIf.Cells(statRow, 2).Font.Color = RGB(89, 89, 89)

        wsWhatIf.Cells(statRow, 3).Value = statsData(k + 1)
        wsWhatIf.Cells(statRow, 3).Font.Bold = True
        wsWhatIf.Cells(statRow, 3).Font.Size = 10
        wsWhatIf.Cells(statRow, 3).Font.Color = RGB(31, 56, 100)
        wsWhatIf.Cells(statRow, 3).Interior.Color = RGB(226, 239, 218)

        If statsData(k + 2) <> "" Then
            wsWhatIf.Cells(statRow, 4).Value = statsData(k + 2)
            wsWhatIf.Cells(statRow, 4).Font.Size = 9
            wsWhatIf.Cells(statRow, 4).Font.Italic = True
            wsWhatIf.Cells(statRow, 4).Font.Color = RGB(89, 89, 89)
        End If
        statRow = statRow + 1
    Next k

    ' ── 5. Scenario Comparison ────────────────────
    statRow = statRow + 2
    wsWhatIf.Cells(statRow, 2).Value = "SCENARIO COMPARISON"
    wsWhatIf.Cells(statRow, 2).Font.Bold = True
    wsWhatIf.Cells(statRow, 2).Font.Size = 11
    wsWhatIf.Cells(statRow, 2).Font.Color = RGB(31, 56, 100)

    statRow = statRow + 1
    Dim scenarios(5, 2) As String
    scenarios(0, 0) = "Scenario":   scenarios(0, 1) = "YTM":    scenarios(0, 2) = "Mod Duration"
    scenarios(1, 0) = "Bull Case": scenarios(1, 1) = "-100bp": scenarios(1, 2) = ""
    scenarios(2, 0) = "Base Case": scenarios(2, 1) = "8.75%":  scenarios(2, 2) = ""
    scenarios(3, 0) = "Bear Case": scenarios(3, 1) = "+100bp": scenarios(3, 2) = ""
    scenarios(4, 0) = "Stress +300bp": scenarios(4, 1) = "+300bp": scenarios(4, 2) = ""
    scenarios(5, 0) = "Stress -300bp": scenarios(5, 1) = "-300bp": scenarios(5, 2) = ""

    For i = 0 To 5
        wsWhatIf.Cells(statRow + i, 2).Value = scenarios(i, 0)
        wsWhatIf.Cells(statRow + i, 2).Font.Size = 10

        wsWhatIf.Cells(statRow + i, 3).Value = scenarios(i, 1)
        wsWhatIf.Cells(statRow + i, 3).Font.Size = 10
        wsWhatIf.Cells(statRow + i, 3).Font.Bold = True
    Next i

    MsgBox "Simulation complete in " & Round(endTime - startTime, 1) & _
           " seconds." & Chr(13) & Chr(10) & _
           "VaR (95%): ₹" & Round(meanPrice - var95, 2) & " Cr" & Chr(13) & Chr(10) & _
           "CVaR (99%): ₹" & Round(meanPrice - cvar99, 2) & " Cr", _
           vbInformation, "Bolt 🦞 Duration Model"

End Sub

' ──────────────────────────────────────────────
' Quick bond pricing function
' ──────────────────────────────────────────────
Function BondPrice(principal As Double, couponRate As Double, _
                   ytm As Double, tenor As Double, freq As Double) As Double
    Dim n As Long
    n = Int(tenor * freq)
    Dim cf As Double, pv As Double
    Dim p As Long
    pv = 0
    For p = 1 To n
        cf = principal * couponRate / freq
        If p = n Then cf = cf + principal
        pv = pv + cf / (1 + ytm / freq) ^ p
    Next p
    BondPrice = pv
End Function

' ──────────────────────────────────────────────
' Modified Duration
' ──────────────────────────────────────────────
Function ModDuration(principal As Double, couponRate As Double, _
                     ytm As Double, tenor As Double, freq As Double) As Double
    Dim n As Long
    n = Int(tenor * freq)
    Dim macDur As Double
    Dim t As Long, pv As Double, totalPV As Double
    Dim weighted As Double

    totalPV = 0
    weighted = 0
    For t = 1 To n
        pv = (principal * couponRate / freq) / (1 + ytm / freq) ^ t
        If t = n Then pv = pv + principal / (1 + ytm / freq) ^ t
        totalPV = totalPV + pv
        weighted = weighted + (t / freq) * pv
    Next t

    If totalPV > 0 Then
        macDur = weighted / totalPV
        ModDuration = macDur / (1 + ytm / freq)
    End If
End Function

' ──────────────────────────────────────────────
' Scenario builder — generates full scenario table
' ──────────────────────────────────────────────
Sub RunScenarioAnalysis()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("What-If")

    Dim ytm As Double, tenor As Double
    Dim principal As Double, couponRate As Double, freq As Double

    principal  = ThisWorkbook.Sheets("Inputs").Range("C6").Value
    couponRate = ThisWorkbook.Sheets("Inputs").Range("C7").Value
    ytm        = ThisWorkbook.Sheets("Inputs").Range("C9").Value
    tenor      = ThisWorkbook.Sheets("Inputs").Range("C10").Value
    freq       = ThisWorkbook.Sheets("Inputs").Range("C8").Value

    Dim scenarios(6, 5) As Double
    Dim yields(6) As Double, prices(6) As Double
    Dim md As Double

    yields(0) = ytm - 0.03
    yields(1) = ytm - 0.02
    yields(2) = ytm - 0.01
    yields(3) = ytm
    yields(4) = ytm + 0.01
    yields(5) = ytm + 0.02
    yields(6) = ytm + 0.03

    Dim i As Integer
    For i = 0 To 6
        prices(i) = BondPrice(principal, couponRate, yields(i), tenor, freq)
        md = ModDuration(principal, couponRate, yields(i), tenor, freq)

        ' Write to sheet starting row 50
        ws.Cells(50 + i, 2).Value = "YTM " & Format(yields(i), "0.00%")
        ws.Cells(50 + i, 3).Value = Round(md, 4)
        ws.Cells(50 + i, 4).Value = Round(prices(i), 2)
        ws.Cells(50 + i, 5).Value = Round(prices(i) - prices(3), 2)
    Next i

    MsgBox "Scenario analysis complete — see rows 50+ on What-If sheet", _
           vbInformation, "Bolt 🦞"
End Sub
