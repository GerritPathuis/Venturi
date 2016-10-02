Imports System.Text
Imports System.IO
Imports System.Configuration
Imports System.Math
Imports System.Collections.Generic
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word

Public Class Form1
    Dim flow_kghr, flow_kgs, flow_m3sec As Double
    Dim dia_in, dia_keel, beta As Double                'Dimensions
    Dim kin_visco, dyn_visco, density As Double         'Medium info
    Dim C_classic, Reynolds, area_in, speed_inlet As Double   'Venturi data
    Dim p1_tap, p2_tap, dp_tap, kappa, tou As Double    'Pressures
    Dim dp_venturi, zeta As Double
    Dim exp_factor, exp_factor1, exp_factor2, exp_factor3 As Double
    Dim A2a, A2b, a2c As Double

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox21.Text =
       "ISO5167-1:2003" & vbCrLf &
       "ISO5167-4:2003" & vbCrLf &
       "Classieke Venturi diameter 200-1200mm"

        '------------- Initial values----------------------
        flow_kghr = 30000           '[kg/m3]
        flow_kgs = flow_kghr / 3600 '[kg/sec]
        density = 1.2               '[kg/m3]
        kappa = 1.4                 'Isentropic exponent
        kin_visco = 15.1 * 10 ^ -6  '[m2/s]
        p1_tap = 101325             '[pa]
        dp_tap = 300                '[pa]
        dia_in = 0.8                '[m] classis venturi inlet diameter = outlet diameter
        beta = 0.5                  '[-]
        C_classic = 0.985           'See ISO5167-4 section 5.5.4

        '--------- calc ---------------
        flow_m3sec = flow_kghr / (3600 * density)   '[m3/s]
        area_in = Math.PI / 4 * dia_in ^ 2          '[m2]
        speed_inlet = flow_m3sec / area_in                '[m/s] keel
        p2_tap = p1_tap - dp_tap
        tou = p2_tap / p1_tap                       'Pressure ratio
        dia_keel = beta * dia_in
        dyn_visco = kin_visco / density             'Calc dyn visco

        '----------- terug zetten op het scherm-------------
        present_results()
        Button1.PerformClick()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown2.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown5.ValueChanged
        Dim Ecc1, Ecc2, Ecc3 As Double
        Dim Dev1, Dev2, Dev3 As Double

        get_data_from_screen()

        Ecc1 = 0        'Start lower limit of eccentricity [-]
        Ecc2 = 1.0      'Start upper limit of eccentricity [-]
        Ecc3 = 0.5      'In the middle of eccentricity [-]

        Dev1 = calc_A2(Ecc1)
        Dev2 = calc_A2(Ecc2)
        Dev3 = calc_A2(Ecc3)

        '-------------Iteratie 30x halveren moet voldoende zijn ---------------
        '---------- Exc= excentricity, looking for Deviation is zero ---------

        For jjr = 0 To 30
            If Dev1 * Dev3 < 0 Then
                Ecc2 = Ecc3
            Else
                Ecc1 = Ecc3
            End If
            Ecc3 = (Ecc1 + Ecc2) / 2
            Dev1 = calc_A2(Ecc1)
            Dev2 = calc_A2(Ecc2)
            Dev3 = calc_A2(Ecc3)
        Next jjr
        beta = Ecc3
        dia_keel = beta * dia_in

        '-------- Controle nulpunt zoek functie ----------------
        If Dev3 > 0.00001 Then
            GroupBox4.BackColor = Color.Red
        Else
            GroupBox4.BackColor = Color.Transparent
        End If

        '-------- Unrecovered pressure loss over the complete venturi assembly----
        'dp_venturi = 0.15 * dp_tap
        dp_venturi = (-0.017 * beta + 0.191) * dp_tap

        '--------- resistance coefficient venturi assembly 
        zeta = 2 * dp_venturi / (density * speed_inlet ^ 2)

        draw_chart1()
        present_results()
    End Sub

    Private Function calc_A2(betaa As Double)

        '----- calc -------------
        p2_tap = p1_tap - dp_tap
        tou = p2_tap / p1_tap                       'Pressure ratio

        '---------- expansie factor ISI 5167-4 Equation 2---------
        exp_factor1 = kappa * tou ^ (2 / kappa)
        exp_factor1 /= kappa - 1

        exp_factor2 = 1 - betaa ^ 4
        exp_factor2 /= 1 - betaa ^ 4 * tou ^ (2 / kappa)

        exp_factor3 = 1 - tou ^ ((kappa - 1) / kappa)
        exp_factor3 /= 1 - tou

        exp_factor = Math.Sqrt(exp_factor1 * exp_factor2 * exp_factor3)

        '------------- itteratie-------------------
        flow_kghr = NumericUpDown1.Value            '[kg/h]
        flow_kgs = flow_kghr / 3600                 '[kg/sec]
        flow_m3sec = flow_kghr / (3600 * density)   '[m3/s]

        area_in = Math.PI / 4 * dia_in ^ 2          '[m2]
        speed_inlet = flow_m3sec / area_in          '[m/s] inlet

        Reynolds = speed_inlet * dia_in * density / kin_visco

        '------- ISO5167-1:2003, SECTION 5.2 page 8-------------
        A2b = C_classic * exp_factor * betaa ^ 2 / Math.Sqrt(1 - betaa ^ 4)
        A2a = 4 * flow_kgs / (PI * dia_in ^ 2 * Math.Sqrt(2 * dp_tap * density))

        a2c = A2a - A2b
        Return (a2c)
    End Function
    Private Sub present_results()
        Try
            NumericUpDown1.Value = flow_kghr            '[kg/m3]
            NumericUpDown7.Value = kappa                'Isentropic exponent
            NumericUpDown2.Value = density              '[kg/m3]
            NumericUpDown6.Value = kin_visco * 10 ^ 6   'kin_visco
            NumericUpDown11.Value = p1_tap / 100        '[mBar]->[pa]
            NumericUpDown8.Value = dp_tap / 100         '[mBar]->[pa]
            NumericUpDown4.Value = dia_in * 1000        '[m] classis venturi inlet diameter = outlet diameter
            NumericUpDown5.Value = beta                 '[-]

            TextBox1.Text = Math.Round(dia_keel * 1000, 0).ToString     '[mm] keel diameter
            TextBox2.Text = C_classic.ToString
            TextBox3.Text = Math.Round(Reynolds, 0).ToString            '[-]
            TextBox4.Text = Math.Round(speed_inlet, 1).ToString               '[m/s]
            TextBox5.Text = Math.Round(exp_factor, 3).ToString          '[-]
            TextBox13.Text = Math.Round(p2_tap / 100, 1).ToString       '[Pa]->[mBar]
            TextBox12.Text = Math.Round(tou, 4).ToString
            TextBox14.Text = Math.Round(dyn_visco * 10 ^ 6, 2).ToString
            TextBox15.Text = Round(dia_in * 1000, 0).ToString       'Diameter in
            TextBox16.Text = Math.Round(flow_m3sec, 3).ToString
            TextBox17.Text = Round(dia_keel * 1000, 0).ToString     'Diameter keel
            TextBox23.Text = Round(dp_venturi / 100, 2).ToString    'Unrecovered pressure loos [mBar]
            TextBox26.Text = Round(zeta, 2).ToString                'Resistance coeffi venturi assembly

            '------- Beta check --------------
            If beta < 0.4 Or beta > 0.7 Then
                NumericUpDown5.BackColor = Color.Red
            Else
                NumericUpDown5.BackColor = Color.LightGreen
            End If

            '------- Tou check --------------
            If tou < 0.75 Then
                TextBox12.BackColor = Color.Red
            Else
                TextBox12.BackColor = Color.LightGreen
            End If

            '------- Reynolds check-----------
            If Reynolds < 2.0 * 10 ^ 5 Or Reynolds > 2.0 * 10 ^ 6 Then
                TextBox3.BackColor = Color.Red
                If Reynolds < 2.0 * 10 ^ 5 Then Label10.Text = "Reynolds, Te lage snelheid"
                If Reynolds > 2.0 * 10 ^ 6 Then Label10.Text = "Reynolds, Te Hoge snelheid"
            Else
                TextBox3.BackColor = Color.LightGreen
                Label10.Text = "Reynolds OK"
            End If
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 845")  ' Show the exception's message.
        End Try
    End Sub

    Private Sub draw_chart1()
        Dim x, y As Double
        Try
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()
            Chart1.Series.Add("Series0")
            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Titles.Add("Determine Beta" & vbCrLf & "ISO 5167-1:2003, Section 5.2")
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)
            Chart1.Series(0).Name = "Koppel[%]"
            Chart1.Series(0).Color = Color.Blue
            Chart1.Series(0).IsVisibleInLegend = False
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 1
            Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Invariant A2"
            Chart1.ChartAreas("ChartArea0").AxisX.Title = "Beta [-]"

            '------ data for the Chart -----------------------------
            For x = 0 To 1.0 Step 0.01
                y = calc_A2(x)
                Chart1.Series(0).Points.AddXY(x, y)
            Next x

            '------ data for the actual beta value -----------------
            calc_A2(beta)
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 206")  ' Show the exception's message.
        End Try
    End Sub
    '-------------------- Dimension of the Venturi ----------------
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabControl1.Enter, NumericUpDown9.ValueChanged, NumericUpDown3.ValueChanged
        Dim Length(10) As Double
        Dim deltad As Double

        deltad = (dia_in - dia_keel) / 2
        TextBox15.Text = Round(dia_in * 1000, 0).ToString       'Diameter in
        TextBox17.Text = Round(dia_keel * 1000, 0).ToString     'Diameter keel

        Length(0) = 2 * dia_in                                  'Bocht R=D 
        Length(1) = 3 * dia_in                                  'Recht in 
        Length(2) = deltad / Math.Tan(NumericUpDown3.Value * Math.PI / 180)       'Convergeren
        Length(3) = dia_keel                                    'Meten
        Length(4) = deltad / Math.Tan(NumericUpDown9.Value * Math.PI / 180)       'Divergeren
        Length(5) = 3 * dia_in                                  'Recht uit
        Length(6) = dia_in / 4                                  'Lucht inlaat
        Length(7) = dia_in                                      'Chinese hat
        Length(8) = Length(0) + Length(1) + Length(2) + Length(3) + Length(4) + Length(5) + Length(6) + Length(7)

        TextBox20.Text = Round(Length(0) * 1000, 0).ToString
        TextBox6.Text = Round(Length(1) * 1000, 0).ToString
        TextBox7.Text = Round(Length(2) * 1000, 0).ToString
        TextBox8.Text = Round(Length(3) * 1000, 0).ToString
        TextBox9.Text = Round(Length(4) * 1000, 0).ToString
        TextBox10.Text = Round(Length(5) * 1000, 0).ToString
        TextBox18.Text = Round(Length(6) * 1000, 0).ToString
        TextBox19.Text = Round(Length(7) * 1000, 0).ToString
        TextBox11.Text = Round(Length(8) * 1000, 0).ToString

        TextBox22.Text = Round(dia_keel * 1000, 0).ToString     'Length C
    End Sub

    Private Sub get_data_from_screen()
        Try
            flow_kghr = NumericUpDown1.Value            '[kg/m3]
            kappa = NumericUpDown7.Value                'Isentropic exponent
            density = NumericUpDown2.Value              '[kg/m3]
            kin_visco = NumericUpDown6.Value / 10 ^ 6   'kin_visco
            p1_tap = NumericUpDown11.Value * 100        '[mBar]->[pa]
            dp_tap = NumericUpDown8.Value * 100         '[mBar]->[pa]
            dia_in = NumericUpDown4.Value / 1000        '[m] classis venturi inlet diameter = outlet diameter
            dyn_visco = kin_visco / density

        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 254")  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim oWord As Word.Application ' = Nothing

        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara4 As Word.Paragraph
        Dim row, font_sizze As Integer
        Dim ufilename As String

        Try
            oWord = New Word.Application()

            'Start Word and open the document template. 
            font_sizze = 9
            oWord = CreateObject("Word.Application")
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = font_sizze + 3
            oPara1.Range.Font.Bold = True
            oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Range.Font.Size = font_sizze + 1
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = False
            oPara2.Range.Text = "Classical Venturi tube acc ISO5167-1,-4:2003" & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = False
            oTable.Rows.Item(1).Range.Font.Bold = True
            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox24.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Project number "
            oTable.Cell(row, 2).Range.Text = TextBox25.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author "
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date "
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = True
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a 16 x 3 table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 22, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = False
            oTable.Rows.Item(1).Range.Font.Bold = True
            oTable.Rows.Item(1).Range.Font.Size = font_sizze + 2
            row = 1
            oTable.Cell(row, 1).Range.Text = "Input Data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Air density"
            oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Kinematic visco"
            oTable.Cell(row, 2).Range.Text = NumericUpDown6.Value
            oTable.Cell(row, 3).Range.Text = "[m2/sec 10^-6]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Dynamic visco"
            oTable.Cell(row, 2).Range.Text = TextBox14.Text
            oTable.Cell(row, 3).Range.Text = "[Pa.s 10^-6]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Isentropic exponent"
            oTable.Cell(row, 2).Range.Text = NumericUpDown7.Value
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet pressure"
            oTable.Cell(row, 2).Range.Text = NumericUpDown11.Value
            oTable.Cell(row, 3).Range.Text = "[mBar abs]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "dp @ max flow"
            oTable.Cell(row, 2).Range.Text = NumericUpDown8.Value
            oTable.Cell(row, 3).Range.Text = "[mBar]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Mass flow"
            oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value
            oTable.Cell(row, 3).Range.Text = "[kg/h]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Volume flow"
            oTable.Cell(row, 2).Range.Text = TextBox16.Text
            oTable.Cell(row, 3).Range.Text = "[m3/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet diameter"
            oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Throut diameter"
            oTable.Cell(row, 2).Range.Text = TextBox1.Text
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet speed_inlet"
            oTable.Cell(row, 2).Range.Text = TextBox4.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Reynolds"
            oTable.Cell(row, 2).Range.Text = TextBox3.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Beta"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown5.Value, 2)
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Discharge Coefficient"
            oTable.Cell(row, 2).Range.Text = TextBox2.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Expansion factor"
            oTable.Cell(row, 2).Range.Text = TextBox5.Text
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Uncovered pressure loss"
            oTable.Cell(row, 2).Range.Text = TextBox23.Text
            oTable.Cell(row, 3).Range.Text = "[mbar]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.4)   'Change width of columns 1 & 2.
            oTable.Columns(2).Width = oWord.InchesToPoints(1.2)
            oTable.Columns(3).Width = oWord.InchesToPoints(1.3)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------save picture ---------------- 
            draw_chart2()
            Chart2.SaveImage("c:\Temp\MainChart.gif", System.Drawing.Imaging.ImageFormat.Gif)
            oPara4 = oDoc.Content.Paragraphs.Add
            oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            oPara4.Range.InlineShapes.AddPicture("c:\Temp\MainChart.gif")
            oPara4.Range.InlineShapes.Item(1).LockAspectRatio = True
            oPara4.Range.InlineShapes.Item(1).Width = 310
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '--------------Save file word file------------------
            'See https://msdn.microsoft.com/en-us/library/63w57f4b.aspx
            ufilename = "N:\Engineering\VBasic\Rapport_copy\Campbell_report_" & DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss") & ".docx"

            If Directory.Exists("N:\Engineering\VBasic\Rapport_copy") Then
                'GroupBox12.Text = "File saved at " & ufilename
                oWord.ActiveDocument.SaveAs(ufilename)
            End If
        Catch ex As Exception
            MessageBox.Show("Bestaat directory N:\Engineering\VBasic\Rapport_copy\ ? " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter
        draw_chart2()
    End Sub

    Private Sub draw_chart2()
        Dim x, y As Double
        Try
            Chart2.Series.Clear()
            Chart2.ChartAreas.Clear()
            Chart2.Titles.Clear()
            Chart2.Series.Add("Series0")
            Chart2.ChartAreas.Add("ChartArea0")
            Chart2.Series(0).ChartArea = "ChartArea0"
            Chart2.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart2.Titles.Add("Venturi flow computation acc. " & "ISO 5167-4:2003 Chapter 4")
            Chart2.Titles.Add("Discharge Coefficient= " & C_classic.ToString & ", Dia.throat= " & Round(dia_keel * 1000, 1).ToString & " [mm]" & ", Density= " & density.ToString & " [kg/m3]" & ", K= " & kappa.ToString & " [-]")
            Chart2.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)
            Chart2.Series(0).Color = Color.Black
            Chart2.Series(0).IsVisibleInLegend = False
            Chart2.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart2.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
            Chart2.ChartAreas("ChartArea0").AxisY.Title = "Flow [kg/hr]"
            Chart2.ChartAreas("ChartArea0").AxisX.Title = "dp_tap [Pa]"

            '----------------- data for the Chart -----------------
            '--------------- see ISO 5167-4 Equation 1-------------
            For x = 0 To dp_tap Step 1
                y = C_classic / Sqrt(1 - beta ^ 4)
                y *= exp_factor * PI / 4 * dia_keel ^ 2 * Sqrt(2 * x * density)
                y *= 3600                               '[kg/h]
                Chart2.Series(0).Points.AddXY(x, y)
            Next x
        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 476")  ' Show the exception's message.
        End Try
    End Sub
End Class
