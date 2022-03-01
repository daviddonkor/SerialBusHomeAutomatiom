Imports System.Runtime.InteropServices
Imports System.Threading

Public Class Form1

    <DllImport("InpOut32.dll", CharSet:=CharSet.Auto, EntryPoint:="Inp32")> _
    Shared Function Inp32(ByVal PortAddress As Short) As Short
    End Function

    <DllImport("InpOut32.dll", CharSet:=CharSet.Auto, EntryPoint:="Out32")> _
    Shared Sub Out32(ByVal PortAddress As Short, ByVal Data As Short)
    End Sub

    <DllImport("Inpout32.dll", CharSet:=CharSet.Auto, EntryPoint:="IsInpOutDriverOpen")> _
    Shared Function IsInpOutDriverOpen() As UInt32
    End Function

    <DllImport("Inpoutx64.dll", CharSet:=CharSet.Auto, EntryPoint:="Inp32")> _
    Shared Function Inp32_x64(ByVal PortAddress As Short) As Short
    End Function

    <DllImport("Inpoutx64.dll", CharSet:=CharSet.Auto, EntryPoint:="Out32")> _
    Shared Sub Out32_x64(ByVal PortAddress As Short, ByVal Data As Short)
    End Sub

    <DllImport("Inpoutx64.dll", CharSet:=CharSet.Auto, EntryPoint:="IsInpOutDriverOpen")> _
    Shared Function IsInpOutDriverOpen_x64() As UInt32
    End Function


    Dim m_bX64 As Boolean = False, onbits(8) As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles device2on.Click
        'Try
        '    Dim iPort As Short
        '    iPort = Convert.ToInt16(TextBox1.Text)

        '    If (m_bX64) Then
        '        TextBox2.Text = Inp32_x64(iPort).ToString()
        '    Else
        '        TextBox2.Text = Inp32(iPort).ToString()
        '    End If

        'Catch ex As Exception
        '    MessageBox.Show("An error occured:\n" + ex.Message)
        'End Try


        Try
            Dim iPort As Short
            Dim iData As Short, bits As Integer
            onbits(1) = 2

            For i = 0 To onbits.Length - 1
                bits += onbits(i)
            Next
            iPort = Convert.ToInt16("888")
            iData = Convert.ToInt16(bits.ToString)

            If (m_bX64) Then
                Out32_x64(iPort, iData)
            Else
                Out32(iPort, iData)
            End If
            device2status.Text = "On"
        Catch ex As Exception
            MessageBox.Show("An error occured:\n" + ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles device1on .Click 
        Try
            Dim iPort As Short
            Dim iData As Short, bits As Integer
            onbits(0) = 1

            For i = 0 To onbits.Length - 1
                bits += onbits(i)
            Next
            iPort = Convert.ToInt16("888")
            iData = Convert.ToInt16(bits.ToString)


            If (m_bX64) Then
                Out32_x64(iPort, iData)
            Else
                Out32(iPort, iData)
            End If
            device1status.Text = "On"
        Catch ex As Exception
            MessageBox.Show("An error occured:\n" + ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim nResult As UInt32

        onbits = {0, 0, 0, 0, 0, 0, 0}

        Try
            nResult = IsInpOutDriverOpen()
        Catch ex As Exception
            nResult = IsInpOutDriverOpen_x64()
            If (nResult <> 0) Then
                m_bX64 = True
            End If
        End Try

        If (nResult = 0) Then
            Label1.Text = "Unable to open InpOut driver"
        End If
    End Sub

    Private Sub Beep(ByVal freq As UInt32)
        If (m_bX64) Then
            Out32_x64(&H43, &HB6)
            Out32_x64(&H42, (freq And &HFF))
            Out32_x64(&H42, (freq >> 9))
            System.Threading.Thread.Sleep(10)
            Out32_x64(&H61, (Convert.ToByte(Inp32_x64(&H61)) Or &H3))
        Else
            Out32(&H43, &HB6)
            Out32(&H42, (freq And &HFF))
            Out32(&H42, (freq >> 9))
            System.Threading.Thread.Sleep(10)
            Out32(&H61, (Convert.ToByte(Inp32(&H61)) Or &H3))
        End If
    End Sub

    Private Sub StopBeep()
        If (m_bX64) Then
            Out32_x64(&H61, (Convert.ToByte(Inp32_x64(&H61)) And &HFC))
        Else
            Out32(&H61, (Convert.ToByte(Inp32(&H61)) And &HFC))
        End If
    End Sub

    Private Sub ThreadBeeper()
        Dim i As UInteger
        For i = 440000 To 500000 Step 1000
            Dim freq As UInteger = 1193180000 / i '440Hz
            Beep(freq)
        Next i
        StopBeep()
    End Sub


    Private Sub button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles device7on.Click, device3off.Click, device3on.Click, device4off.Click, device4on.Click, device5off.Click, device5on.Click, device6off.Click, device6on.Click, device7off.Click, device8off.Click, device8on.Click
        Dim t As New System.Threading.Thread(New System.Threading.ThreadStart(AddressOf ThreadBeeper))
        t.Start()
    End Sub

    Private Sub device1off_Click(sender As Object, e As EventArgs) Handles device1off.Click
        Try
            Dim iPort As Short
            Dim iData As Short, bits As Integer

            iPort = Convert.ToInt16("888")
            onbits(0) = 0
            For i = 0 To onbits.Length - 1
                bits += onbits(i)
            Next
            iData = Convert.ToInt16(bits.ToString)

            If (m_bX64) Then
                Out32_x64(iPort, iData)
            Else
                Out32(iPort, iData)
            End If
            device1status.Text = "Off"
        Catch ex As Exception
            MessageBox.Show("An error occured:\n" + ex.Message)
        End Try
    End Sub

    Private Sub device2off_Click(sender As Object, e As EventArgs) Handles device2off.Click
        Try
            Dim iPort As Short
            Dim iData As Short, bits As Integer
            onbits(1) = 0
            For i = 0 To onbits.Length - 1
                bits += onbits(i)
            Next
            iPort = Convert.ToInt16("888")
            iData = Convert.ToInt16(bits.ToString)

            If (m_bX64) Then
                Out32_x64(iPort, iData)
            Else
                Out32(iPort, iData)
            End If
            device2status.Text = "Off"
        Catch ex As Exception
            MessageBox.Show("An error occured:\n" + ex.Message)
        End Try
    End Sub
End Class
