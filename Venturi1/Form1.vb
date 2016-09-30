Imports System.Text
Imports System.IO
Imports System.Configuration
Imports System.Math
Imports System.Collections.Generic
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
'Imports Word = Microsoft.Office.Interop.Word

Public Class Form1
    Dim flow_kghr, flow_m3sec As Double
    Dim dia_in, dia_keel, beta As Double                'Dimensions
    Dim kin_visco, dyn_visco, density As Double         'Medium info
    Dim C_classic, Reynolds, area_in, speed As Double   'Venturi data
    Dim p1_tap, p2_tap, dp_tap, kappa, tou As Double    'Pressures
    Dim exp_factor, exp_factor1, exp_factor2, exp_factor3 As Double
    Dim A2a, A2b, a2c As Double

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Start bbb
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown2.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown5.ValueChanged
        Dim Ecc1, Ecc2, Ecc3 As Double
        Dim Dev1, Dev2, Dev3 As Double

        beta = NumericUpDown5.Value
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
        NumericUpDown5.Value = Round(Ecc3, 3).ToString              'Beta diameter ratio
        beta = Ecc3
        dia_keel = beta * dia_in

        '-------- Controle nulpunt zoek functie ----------------
        If Dev3 > 0.01 Then
            GroupBox4.BackColor = Color.Red
        Else
            GroupBox4.BackColor = Color.Transparent
        End If

        present_results()
        draw_chart1()
    End Sub

    Private Function calc_A2(betaa As Double)
        C_classic = 0.985                           'See ISO5167-4 section 5.5.4
        kappa = NumericUpDown7.Value                'Isentropic exponent
        density = NumericUpDown2.Value              '[kg/m3]
        kin_visco = NumericUpDown6.Value * 10 ^ -6
        p1_tap = NumericUpDown11.Value * 100        '[mBar]->[pa]
        dp_tap = NumericUpDown8.Value * 100         '[mBar]->[pa]
        dia_in = NumericUpDown4.Value / 1000        '[m] classis venturi inlet diameter = outlet diameter
        dia_keel = betaa * dia_in                    '[m]

        '-------------VB Prevent problems ----------------
        If density = 0 Then density = 1             'Prevent problems
        dyn_visco = kin_visco / density

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
        flow_kghr = NumericUpDown1.Value
        flow_m3sec = flow_kghr / (3600 * density)   '[m3/s]

        area_in = Math.PI / 4 * dia_in ^ 2          '[m2]
        speed = flow_m3sec / area_in                '[m/s] keel

        Reynolds = speed * dia_in * density / kin_visco

        A2a = dyn_visco * Reynolds / (dia_in * Math.Sqrt(2 * dp_tap * density))
        A2b = C_classic * exp_factor * betaa ^ 2 / Math.Sqrt(1 - betaa ^ 4)
        a2c = A2a - A2b
        Return (a2c)
    End Function
    Private Sub present_results()
        TextBox1.Text = Math.Round(dia_keel * 1000, 0).ToString     '[mm] keel diameter
        TextBox2.Text = C_classic.ToString
        TextBox3.Text = Math.Round(Reynolds, 0).ToString            '[-]
        TextBox4.Text = Math.Round(speed, 1).ToString               '[m/s]
        TextBox5.Text = Math.Round(exp_factor, 3).ToString          '[-]
        TextBox13.Text = Math.Round(p2_tap / 100, 1).ToString       '[Pa]->[mBar]
        TextBox12.Text = Math.Round(tou, 4).ToString
        TextBox14.Text = Math.Round(dyn_visco * 10 ^ 6, 2).ToString
        TextBox16.Text = Math.Round(flow_m3sec, 4).ToString

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

    End Sub

    Private Sub draw_chart1()
        Dim x, y As Double
        Try
            'Clear all series And chart areas so we can re-add them
            Chart1.Series.Clear()
            Chart1.ChartAreas.Clear()
            Chart1.Titles.Clear()
            Chart1.Series.Add("Series0")
            Chart1.ChartAreas.Add("ChartArea0")
            Chart1.Series(0).ChartArea = "ChartArea0"
            Chart1.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            Chart1.Titles.Add("Determine Beta" & vbCrLf & "ISO 5167, A2 page 20")
            Chart1.Titles(0).Font = New Font("Arial", 12, System.Drawing.FontStyle.Bold)
            Chart1.Series(0).Name = "Koppel[%]"
            Chart1.Series(0).Color = Color.Blue
            Chart1.Series(0).IsVisibleInLegend = False
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 1
            Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            Chart1.ChartAreas("ChartArea0").AxisY.Title = "Invariant A2"
            Chart1.ChartAreas("ChartArea0").AxisX.Title = "Beta [-]"

            For x = 0 To 1.01 Step 0.01
                y = calc_A2(x)
                Chart1.Series(0).Points.AddXY(x, y)
            Next x

        Catch ex As Exception
            'MessageBox.Show(ex.Message &"Error 845")  ' Show the exception's message.
        End Try
    End Sub
End Class
