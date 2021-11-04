Public Class Form1
    Dim jogador1, jogador2, vencedor As String
    Dim vez As Integer
    Dim mg1, mg2 As Image

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        jogador1 = InputBox("Insira o seu nome..... Você será o X", "Olá Jogador 1 ;)")
        jogador2 = InputBox("Insira o seu nome..... Você será a O", "Olá Jogador 2 ;)")
        Button1.Visible = False
        Label1.Visible = True
        Label2.Visible = True
        Label2.Text = jogador1


        PictureBox1.Enabled = True
        PictureBox2.Enabled = True
        PictureBox3.Enabled = True
        PictureBox4.Enabled = True
        PictureBox5.Enabled = True
        PictureBox6.Enabled = True
        PictureBox7.Enabled = True
        PictureBox8.Enabled = True
        PictureBox9.Enabled = True

    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        vez = vez + 1
        PictureBox1.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If


        If vez Mod 2 = 0 Then
            PictureBox1.Image = mg2
        Else
            PictureBox1.Image = mg1

        End If

        Verificar(jogador1, jogador2)

    End Sub
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        vez = vez + 1
        PictureBox2.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox2.Image = mg2
        Else
            PictureBox2.Image = mg1

        End If

        Verificar(jogador1, jogador2)

    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click

        vez = vez + 1
        PictureBox3.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox3.Image = mg2
        Else
            PictureBox3.Image = mg1
        End If

        Verificar(jogador1, jogador2)

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        vez = vez + 1
        PictureBox4.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox4.Image = mg2
        Else
            PictureBox4.Image = mg1

        End If

        Verificar(jogador1, jogador2)

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click

        vez = vez + 1
        PictureBox5.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox5.Image = mg2
        Else
            PictureBox5.Image = mg1

        End If

        Verificar(jogador1, jogador2)

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        vez = vez + 1
        PictureBox6.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox6.Image = mg2
        Else
            PictureBox6.Image = mg1

        End If

        Verificar(jogador1, jogador2)

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click

        vez = vez + 1
        PictureBox7.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox7.Image = mg2
        Else
            PictureBox7.Image = mg1

        End If

        Verificar(jogador1, jogador2)
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        vez = vez + 1
        PictureBox8.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2

        End If

        If vez Mod 2 = 0 Then
            PictureBox8.Image = mg2
        Else
            PictureBox8.Image = mg1

        End If

        Verificar(jogador1, jogador2)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        mg1 = My.Resources.Resource1.X_1

        mg2 = My.Resources.Resource1.O_2
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

        vez = vez + 1
        PictureBox9.Enabled = False
        If vez Mod 2 = 0 Then
            Label2.Text = jogador1
        Else
            Label2.Text = jogador2
        End If

        If vez Mod 2 = 0 Then
            PictureBox9.Image = mg2
        Else
            PictureBox9.Image = mg1

        End If
        Verificar(jogador1, jogador2)
    End Sub


    Sub Verificar(ByVal jogador1 As String, ByVal Jogador2 As String)

        If (vez + 1) Mod 2 = 0 Then
            vencedor = jogador1
        Else
            vencedor = Jogador2
        End If


        If (PictureBox1.Image.Equals(PictureBox2.Image) = True) And (PictureBox2.Image.Equals(PictureBox3.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox4.Image.Equals(PictureBox5.Image) = True) And (PictureBox5.Image.Equals(PictureBox6.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox7.Image.Equals(PictureBox8.Image) = True) And (PictureBox8.Image.Equals(PictureBox9.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox1.Image.Equals(PictureBox4.Image) = True) And (PictureBox4.Image.Equals(PictureBox7.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox2.Image.Equals(PictureBox5.Image) = True) And (PictureBox5.Image.Equals(PictureBox8.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox3.Image.Equals(PictureBox6.Image) = True) And (PictureBox6.Image.Equals(PictureBox9.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox1.Image.Equals(PictureBox5.Image) = True) And (PictureBox5.Image.Equals(PictureBox9.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf (PictureBox3.Image.Equals(PictureBox5.Image) = True) And (PictureBox5.Image.Equals(PictureBox7.Image) = True) Then
            MsgBox("Parabens " & vencedor & " você ganhou")

        ElseIf vez > 8 Then
            MsgBox(" Empate ")
        End If

    End Sub
End Class