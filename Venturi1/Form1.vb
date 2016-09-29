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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown2.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown5.ValueChanged
        Dim flow_kghr, flow_m3sec, density As Double
        Dim dia_in, dia_keel, beta, C_classic, Reynolds, kin_visco, dyn_visco, area_in, speed As Double
        Dim p1_tap, p2_tap, dp_tap, kappa, tou As Double
        Dim exp_factor, exp_factor1, exp_factor2, exp_factor3 As Double
        Dim A2a, A2b, a2c As Double

        C_classic = 0.985                           'See ISO5167-4 section 5.5.4

        kappa = NumericUpDown7.Value                'Isentropic exponent
        density = NumericUpDown2.Value              '[kg/m3]
        kin_visco = NumericUpDown6.Value * 10 ^ -6
        p1_tap = NumericUpDown11.Value * 100        '[mBar]->[pa]
        dp_tap = NumericUpDown8.Value * 100         '[mBar]->[pa]

        dia_in = NumericUpDown4.Value / 1000        '[m] classis venturi inlet diameter = outlet diameter
        beta = NumericUpDown5.Value  '[-]
        dia_keel = beta * dia_in                    '[m]

        '-------------VB Prevent problems ----------------
        If density = 0 Then density = 1             'Prevent problems
        dyn_visco = kin_visco / density

        '----- calc -------------
        p2_tap = p1_tap - dp_tap
        tou = p2_tap / p1_tap                       'Pressure ratio


        '---------- expansie factor ISI 5167-4 Equation 2---------
        exp_factor1 = kappa * tou ^ (2 / kappa)
        exp_factor1 /= kappa - 1

        exp_factor2 = 1 - beta ^ 4
        exp_factor2 /= 1 - beta ^ 4 * tou ^ (2 / kappa)

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
        A2b = C_classic * exp_factor * beta ^ 2 / Math.Sqrt(1 - beta ^ 4)
        a2c = A2a - A2b

        TextBox15.Text = a2c.ToString

        '------------- snelheid -------------------
        TextBox1.Text = Math.Round(dia_keel * 1000, 2).ToString     '[mm]
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
End Class
