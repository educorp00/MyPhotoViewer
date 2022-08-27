Imports System.IO
Imports Shell32

Public Class Form1

    ' フルスクリーン・モードかどうかのフラグ
    Private _bScreenMode As Boolean
    ' フルスクリーン表示前のウィンドウの状態を保存する
    Private prevFormState As FormWindowState
    ' 通常表示時のフォームの境界線スタイルを保存する
    Private prevFormStyle As FormBorderStyle
    ' 通常表示時のウィンドウのサイズを保存する
    Private prevFormSize As Size

    Private _pictfilename As String

    Private db_files As String()
    Private db_num As Integer()

    Private slideshow_on As Boolean = False
    Private media_pict As Boolean = True
    Private metext As String

    Private exts As String() = {".jpeg", ".jpg", ".bmp", ".png", ".gif", ".mp4", ".avi", ".mov", ".mpg", ".wmv", ".m4v"}

    '表示する画像
    Private currentImage As Image
    '倍率
    Private zoomRatio As Double = 1.0
    Private initzoomRatio As Double

    '倍率変更後の画像のサイズと位置
    Private drawRectangle As Rectangle
    Private initRectangle As Rectangle
    Private mousexy As System.Drawing.Point
    Private dragxy_start As System.Drawing.Point

    Private drag_st As Boolean = False

    Class pictdb
        Public time As DateTime
        Public file As String
        Sub New(time As DateTime, file As String)
            Me.time = time
            Me.file = file


        End Sub

    End Class


    Private Sub Button2_Click(sender As Object, e As EventArgs) 
        If (_bScreenMode = False) Then
            Call Change_fullSCR()
        End If

    End Sub




    Private Sub Change_fullSCR()

        If (_bScreenMode = False) Then
            ' ＜フルスクリーン表示への切り替え処理＞

            TableLayoutPanel1.RowStyles(1) = New RowStyle(SizeType.Absolute, 0)
            Me.BackColor = Color.FromArgb(0, 0, 0)
            AxWindowsMediaPlayer1.uiMode = "none"
            If media_pict = True Then
                Timer1.Enabled = True
                Timer1.Start()
            End If
            slideshow_on = True
            AxWindowsMediaPlayer1.Enabled = False
            AxWindowsMediaPlayer1.fullScreen = False

            AxWindowsMediaPlayer1.Ctlenabled = False

            AxWindowsMediaPlayer1.enableContextMenu = False

            ' ウィンドウの状態を保存する
            prevFormState = Me.WindowState
            ' 境界線スタイルを保存する
            prevFormStyle = Me.FormBorderStyle

            ' 0. 「最大化表示」→「フルスクリーン表示」では
            ' タスク・バーが消えないので、いったん「通常表示」を行う
            If Me.WindowState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
            End If

            ' フォームのサイズを保存する
            prevFormSize = Me.ClientSize

            ' 1. フォームの境界線スタイルを「None」にする
            Me.FormBorderStyle = FormBorderStyle.None
            ' 2. フォームのウィンドウ状態を「最大化」する
            Me.WindowState = FormWindowState.Maximized

            ' フルスクリーン・モードをONにする
            _bScreenMode = True
            If media_pict = False Then
                mediashow(_pictfilename)
            End If

        Else
            ' ＜通常表示／最大化表示への切り替え処理＞
            Timer1.Enabled = False
            Timer1.Stop()
            slideshow_on = False
            TableLayoutPanel1.RowStyles(1) = New RowStyle(SizeType.Absolute, 52)
            Me.BackColor = Color.FromArgb(238, 243, 250)


            ' フォームのウィンドウのサイズを元に戻す
            Me.ClientSize = prevFormSize

            ' 0. 最大化に戻す場合にはいったん通常表示を行う
            ' （フルスクリーン表示の処理とのバランスと取るため）
            If prevFormState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
            End If

            ' 1. フォームの境界線スタイルを元に戻す
            Me.FormBorderStyle = prevFormStyle

            ' 2. フォームのウィンドウ状態を元に戻す
            Me.WindowState = prevFormState


            ' フルスクリーン・モードをOFFにする
            _bScreenMode = False


            AxWindowsMediaPlayer1.uiMode = "full"

            AxWindowsMediaPlayer1.Ctlenabled = True

            AxWindowsMediaPlayer1.enableContextMenu = True

        End If


    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _bScreenMode = False
        'PictureBox1.Controls.Add(TableLayoutPanel1)
        'コマンドライン引数を配列で取得する
        Dim cmds As String() = System.Environment.GetCommandLineArgs()
        'コマンドライン引数を列挙する

        If (cmds.Length > 1) Then
            _pictfilename = cmds(1)

            Dim fileext As String = LCase(System.IO.Path.GetExtension(_pictfilename))
            If exts.Contains(fileext) Then
                Call mediashow(_pictfilename)
                Call builddata(_pictfilename)

            End If
        End If


    End Sub

    Private Sub mediashow(pictfilename As String)
        Console.WriteLine(pictfilename)


        If slideshow_on = True Then
            Timer1.Enabled = False
            Timer1.Stop()
        End If

        media_pict = True
        Dim fileext As String
        fileext = LCase(System.IO.Path.GetExtension(pictfilename))
        If {".jpeg", ".jpg", ".bmp", ".png", ".gif"}.Contains(fileext) Then


            AxWindowsMediaPlayer1.Visible = False
            AxWindowsMediaPlayer1.Ctlcontrols.stop()
            '描画先とするImageオブジェクトを作成する
            Dim canvas As New Bitmap(PictureBox1.Width, PictureBox1.Height)
            'ImageオブジェクトのGraphicsオブジェクトを作成する
            Dim g As Graphics = Graphics.FromImage(canvas)
            Dim fs As FileStream = File.OpenRead(pictfilename)
            '画像ファイルを読み込んで、Imageオブジェクトとして取得する
            currentImage = Image.FromStream(fs, False, False)


            If {".jpeg", ".jpg"}.Contains(fileext) Then

                Dim bmp As New System.Drawing.Bitmap(pictfilename)
                Dim item As System.Drawing.Imaging.PropertyItem

                ' 画像に付与されているEXIF情報を列挙する
                For Each item In bmp.PropertyItems
                    If item.Id <> &H112 Then Continue For

                    ' IFD0 0x0112; Orientationの値を調べる
                    Select Case item.Value(0)
                        Case 3
                            ' 時計回りに180度回転しているので、180度回転して戻す
                            currentImage.RotateFlip(RotateFlipType.Rotate180FlipNone)
                        Case 6
                            ' 時計回りに270度回転しているので、90度回転して戻す
                            currentImage.RotateFlip(RotateFlipType.Rotate90FlipNone)
                        Case 8
                            ' 時計回りに90度回転しているので、270度回転して戻す
                            currentImage.RotateFlip(RotateFlipType.Rotate270FlipNone)
                    End Select


                Next


            End If


            '画像を指定された位置、サイズで描画する
            Dim dispwidth As Integer = PictureBox1.Width
            Dim dispheight As Integer = PictureBox1.Height
            Dim dispx As Integer
            Dim dispy As Integer



            If dispwidth / currentImage.Width < dispheight / currentImage.Height Then
                dispheight = dispwidth / currentImage.Width * currentImage.Height
                dispx = 0
                dispy = (PictureBox1.Height - dispheight) / 2
            Else

                dispwidth = dispheight / currentImage.Height * currentImage.Width
                dispx = (PictureBox1.Width - dispwidth) / 2
                dispy = 0

            End If

            zoomRatio = dispheight / currentImage.Height


            drawRectangle.Width = dispwidth
            drawRectangle.Height = dispheight
            drawRectangle.X = dispx
            drawRectangle.Y = dispy


            zoomRatio = dispwidth / currentImage.Width
            initzoomRatio = zoomRatio




            PictureBox1.Invalidate()

            PictureBox1.Visible = True


            If slideshow_on = True Then
                Timer1.Enabled = True
                Timer1.Start()
            End If



        End If




        If {".mp4", ".avi", ".mov", ".mpg", ".wmv", ".m4v"}.Contains(fileext) Then

            media_pict = False
            If slideshow_on = True Then
                Timer1.Enabled = False
                Timer1.Stop()
            End If

            AxWindowsMediaPlayer1.Width = PictureBox1.Size.Width


            AxWindowsMediaPlayer1.Height = PictureBox1.Size.Height





            PictureBox1.Image = Nothing
            PictureBox1.Visible = False


            AxWindowsMediaPlayer1.Visible = True
                AxWindowsMediaPlayer1.settings.autoStart = True
                AxWindowsMediaPlayer1.URL = pictfilename
                AxWindowsMediaPlayer1.stretchToFit = True




            End If


            Dim shell As New ShellClass()
        Dim objFolder As Folder = shell.NameSpace(System.IO.Path.GetDirectoryName(pictfilename))
        Dim objItem As FolderItem
        Dim mediadate As String = ""
        objItem = objFolder.ParseName(System.IO.Path.GetFileName(pictfilename))
        If {".jpeg", ".jpg"}.Contains(fileext) Then
            mediadate = objFolder.GetDetailsOf(objItem, 12)
        End If

        If {".mp4", ".avi", ".mov", ".mpg", ".wmv", ".m4v"}.Contains(fileext) Then
            mediadate = objFolder.GetDetailsOf(objItem, 208)
        End If

        If mediadate = "" Then
            If DateTime.Compare(DateTime.Parse(objFolder.GetDetailsOf(objItem, 3)), DateTime.Parse(objFolder.GetDetailsOf(objItem, 4))) < 0 Then
                mediadate = objFolder.GetDetailsOf(objItem, 3)
            Else
                mediadate = objFolder.GetDetailsOf(objItem, 4)
            End If
        End If

        mediadate = System.Text.RegularExpressions.Regex.Replace(mediadate, "[^0-9,:,/,\s]", "")

        metext = mediadate & "  (" & System.IO.Path.GetFileName(pictfilename) & ") マイ フォト ビューアー"

        Me.Text = metext





    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Dim fileName As String() = CType(e.Data.GetData(DataFormats.FileDrop, False), String())
        _pictfilename = fileName(0)

        Call mediashow(_pictfilename)
        Me.Activate()
        Call builddata(fileName(0))


    End Sub

    Private Sub builddata(filename As String)

        '       Dim exts As String() = {".jpeg", ".jpg", ".bmp", ".png", ".gif", ".mp4", ".avi", ".mov"}
        Dim files As String() =
        System.IO.Directory.GetFiles(System.IO.Path.GetDirectoryName(filename)).Where(Function(f) exts.Contains(LCase(System.IO.Path.GetExtension(f).ToLower))).ToArray
        Array.Sort(files)
        'Array.Reverse(files)
        ReDim db_files(files.Length - 1)
        Console.WriteLine("files.Length:" & files.Length)
        Console.WriteLine("db_files.Length:" & db_files.Length)

        Array.Copy(files, db_files, files.Length)

        Dim pictdb1 As New List(Of pictdb)


        Dim shell As New ShellClass()
        Dim objFolder As Folder = shell.NameSpace(System.IO.Path.GetDirectoryName(filename))
        Dim objItem As FolderItem
        Dim mediadate As String = ""
        Dim j As Integer = 0
        ' Task.Run(Sub()
        For Each f In files
                         Console.WriteLine("db:" & f)
            mediadate = ""
            objItem = objFolder.ParseName(System.IO.Path.GetFileName(f))

                         If {".jpeg", ".jpg"}.Contains(LCase(System.IO.Path.GetExtension(f))) Then
                             mediadate = objFolder.GetDetailsOf(objItem, 12)

                         End If

                         If {".mp4", ".avi", ".mov", ".mpg", ".wmv", ".m4v"}.Contains(LCase(System.IO.Path.GetExtension(f))) Then
                             mediadate = objFolder.GetDetailsOf(objItem, 208)

                         End If

                         If mediadate = "" Then
                             If DateTime.Compare(DateTime.Parse(objFolder.GetDetailsOf(objItem, 3)), DateTime.Parse(objFolder.GetDetailsOf(objItem, 4))) < 0 Then
                                 mediadate = objFolder.GetDetailsOf(objItem, 3)
                             Else
                                 mediadate = objFolder.GetDetailsOf(objItem, 4)
                             End If
                         End If


            pictdb1.Add(New pictdb(DateTime.Parse(System.Text.RegularExpressions.Regex.Replace(mediadate, "[^0-9,:,/,\s]", "")), f))



        Next

        Dim val = pictdb1.OrderBy(Function(k) k.time).ThenBy(Function(k) k.file).ToList()
        Dim i As Integer = 0


                     For Each f In files
                         db_files(i) = val(i).file
                         Console.WriteLine(val(i).time & " " & db_files(i))
                         i = i + 1
                     Next
        '         End Sub)


    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        'データ形式の確認
        If e.Data.GetDataPresent(DataFormats.FileDrop) = False Then
            Return
        End If

        'ドラッグしているファイル／フォルダの取得
        Dim FilePath() As String =
      CType(e.Data.GetData(DataFormats.FileDrop), String())

        For idx As Integer = 0 To FilePath.Length - 1
            If Not System.IO.File.Exists(FilePath(idx)) Then
                Return
            End If
        Next idx

        Dim fileext As String = LCase(System.IO.Path.GetExtension(FilePath(0)))
        If Not exts.Contains(fileext) Then
            Return
        End If

        'ドロップ可能な場合は、エフェクトを変える
        e.Effect = DragDropEffects.Copy

    End Sub




    Private Sub Form1_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        TableLayoutPanel1.BackColor = Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        TableLayoutPanel1.BackColor = Color.FromArgb(255, 255, 255)
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) 
        Call rev_key()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) 
        Call frw_key()
    End Sub


    Private Sub rev_key()
        If _pictfilename <> "" Then
            Dim fileindex As Integer = Array.IndexOf(db_files, _pictfilename)


            fileindex = fileindex - 1
            If fileindex < 0 Then
                fileindex = db_files.Length - 1
            End If
            Console.WriteLine(fileindex)
            _pictfilename = db_files(fileindex)
            Call mediashow(_pictfilename)
        End If

    End Sub

    Private Sub frw_key()
        If _pictfilename <> "" Then

            Dim fileindex As Integer = Array.IndexOf(db_files, _pictfilename)
            Console.WriteLine("db_files.Length:" & db_files.Count)
            Console.WriteLine("now:" & fileindex)

            fileindex = fileindex + 1
            If fileindex = db_files.Length Then
                fileindex = 0
            End If
            Console.WriteLine(fileindex)

            _pictfilename = db_files(fileindex)
            Call mediashow(_pictfilename)
        End If

    End Sub







    Private Sub AxWindowsMediaPlayer1_PlayStateChange(sender As Object, e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent) Handles AxWindowsMediaPlayer1.PlayStateChange
        Select Case e.newState

            Case 0 ' Undefined


            Case 1 ' Stopped
                If slideshow_on = True Then
                    Timer2.Enabled = True
                End If

            Case 2 ' Paused

            Case 3 ' Playing

            Case 4 ' ScanForward

            Case 5 ' ScanReverse

            Case 6 ' Buffering

            Case 7 ' Waiting

            Case 8 ' MediaEnded

            Case 9 ' Transitioning

            Case 10 ' Ready

            Case 11 ' Reconnecting

            Case 12 ' Last

            Case Else



        End Select

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Call frw_key()
    End Sub



    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) _
        Handles ToolStripMenuItem1.Click,
                ToolStripMenuItem2.Click,
                ToolStripMenuItem3.Click
        'グループのToolStripMenuItemを配列にしておく
        Dim groupMenuItems As ToolStripMenuItem() = New ToolStripMenuItem() {
                Me.ToolStripMenuItem1,
                Me.ToolStripMenuItem2,
                Me.ToolStripMenuItem3}



        'グループのToolStripMenuItemを列挙する
        For Each item As ToolStripMenuItem In groupMenuItems
            If Object.ReferenceEquals(item, sender) Then
                'ClickされたToolStripMenuItemならば、Indeterminateにする
                item.CheckState = CheckState.Indeterminate
            Else
                'ClickされたToolStripMenuItemでなければ、Uncheckedにする
                item.CheckState = CheckState.Unchecked
            End If



        Next


        If Object.ReferenceEquals(Me.ToolStripMenuItem1, sender) Then

            Call setinterval(1)
        End If

        If Object.ReferenceEquals(Me.ToolStripMenuItem2, sender) Then

            Call setinterval(2)
        End If

        If Object.ReferenceEquals(Me.ToolStripMenuItem3, sender) Then
            Call setinterval(3)
        End If



    End Sub

    Private Sub setinterval(intervaltype As Integer)
        Select Case intervaltype
            Case 1
                Timer1.Interval = 7000
            Case 2
                Timer1.Interval = 5000
            Case 3
                Timer1.Interval = 2000
        End Select
    End Sub

    '  Private Sub AxWindowsMediaPlayer1_KeyDownEvent(sender As Object, e As AxWMPLib._WMPOCXEvents_KeyDownEvent) Handles AxWindowsMediaPlayer1.KeyDownEvent
    'Select Case e.nKeyCode
    'Case Keys.Escape
    '
    '    If (Me.WindowState = FormWindowState.Maximized) Then
    '   Me.WindowState = FormWindowState.Normal
    '  End If
    ' If (_bScreenMode = True) Then
    'Call Change_fullSCR()
    'End If
    'Case Keys.Left
    '           Console.WriteLine("右キー")
    'Call rev_key()
    'Case Keys.Right
    '           Console.WriteLine("左キー")
    '
    'Call frw_key()
    '

    '    Case Keys.Enter


    '    Call Change_fullSCR()

    '    End Select
    '   End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        If media_pict = False Then
            Call frw_key()
        End If

        Timer2.Enabled = False

    End Sub

    Private Sub PictureBox1_LoadProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles PictureBox1.LoadProgressChanged
        Me.Text = String.Format("{0}% 読み込みました", e.ProgressPercentage)
        If e.ProgressPercentage = 100 Then
            Call PictureBox1LoadComplete()

        End If
    End Sub

    Private Sub PictureBox1LoadComplete()

    End Sub

    Private Sub PictureBox1_LoadCompleted(sender As Object, e As System.ComponentModel.AsyncCompletedEventArgs) Handles PictureBox1.LoadCompleted
        'Call PictureBox1LoadComplete()
    End Sub



    Private Sub PictureBox_MouseWheel(sender As System.Object,
                             e As MouseEventArgs) Handles PictureBox1.MouseWheel

        Dim pb As PictureBox = DirectCast(sender, PictureBox)
        '        Dim sp0 As System.Drawing.Point = System.Windows.Forms.Cursor.Position
        Dim sp As Point = mousexy


        Dim dispwidth As Integer = pb.Width
        Dim dispheight As Integer = pb.Height
        Dim dispx As Integer
        Dim dispy As Integer

        If dispwidth / currentImage.Width < dispheight / currentImage.Height Then
            dispheight = dispwidth / currentImage.Width * currentImage.Height
            dispx = 0
            dispy = (PictureBox1.Height - dispheight) / 2
        Else

            dispwidth = dispheight / currentImage.Height * currentImage.Width
            dispx = (PictureBox1.Width - dispwidth) / 2
            dispy = 0

        End If
        Dim imgPoint As New Point(CInt(Math.Round((sp.X - drawRectangle.X) / zoomRatio)), CInt(Math.Round((sp.Y - drawRectangle.Y) / zoomRatio)))


        If e.Delta > 0 Then
            zoomRatio *= 2
        Else
            zoomRatio *= 0.5
        End If
        If initzoomRatio > zoomRatio Then



            zoomRatio = dispheight / currentImage.Height


            drawRectangle.Width = dispwidth
            drawRectangle.Height = dispheight
            drawRectangle.X = dispx
            drawRectangle.Y = dispy


            zoomRatio = dispwidth / currentImage.Width
            initzoomRatio = zoomRatio

        Else

            '倍率変更後の画像のサイズと位置を計算する
            drawRectangle.Width = CInt(Math.Round(currentImage.Width * zoomRatio))
            drawRectangle.Height = CInt(Math.Round(currentImage.Height * zoomRatio))
            drawRectangle.X = CInt(Math.Round(sp.X - imgPoint.X * zoomRatio))
            drawRectangle.Y = CInt(Math.Round(sp.Y - imgPoint.Y * zoomRatio))


            If drawRectangle.X > dispx Then
                drawRectangle.X = dispx
            End If

            If drawRectangle.Width + drawRectangle.X < PictureBox1.Width - dispx Then
                drawRectangle.X = PictureBox1.Width - dispx - drawRectangle.Width
            End If


            If drawRectangle.Y > dispy Then
                drawRectangle.Y = dispy
            End If

            If drawRectangle.Height + drawRectangle.Y < PictureBox1.Height - dispy Then
                drawRectangle.Y = PictureBox1.Height - dispy - drawRectangle.Height

            End If






        End If



        PictureBox1.Invalidate()

    End Sub

    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
        If Not (currentImage Is Nothing) Then
            If (zoomRatio > initzoomRatio) Or (currentImage.Width < 320 And currentImage.Height < 320) Then
                e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor
            Else
                e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Default
            End If
            e.Graphics.DrawImage(currentImage, drawRectangle)
        End If
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        mousexy.X = e.X
        mousexy.Y = e.Y
        ' Console.WriteLine(Math.Round((e.X - drawRectangle.X) * zoomRatio) & "," & Math.Round((e.Y - drawRectangle.Y) * zoomRatio))

        If drag_st = True Then

            Dim dispwidth As Integer = PictureBox1.Width
            Dim dispheight As Integer = PictureBox1.Height
            Dim dispx As Integer
            Dim dispy As Integer

            If dispwidth / currentImage.Width < dispheight / currentImage.Height Then
                dispheight = dispwidth / currentImage.Width * currentImage.Height
                dispx = 0
                dispy = (PictureBox1.Height - dispheight) / 2
            Else

                dispwidth = dispheight / currentImage.Height * currentImage.Width
                dispx = (PictureBox1.Width - dispwidth) / 2
                dispy = 0

            End If

            drawRectangle.X = drawRectangle.X + mousexy.X - dragxy_start.X
            drawRectangle.Y = drawRectangle.Y + mousexy.Y - dragxy_start.Y

            If drawRectangle.X > dispx Then
                drawRectangle.X = dispx
            End If

            If drawRectangle.Width + drawRectangle.X < PictureBox1.Width - dispx Then
                drawRectangle.X = PictureBox1.Width - dispx - drawRectangle.Width
            End If


            If drawRectangle.Y > dispy Then
                drawRectangle.Y = dispy
            End If

            If drawRectangle.Height + drawRectangle.Y < PictureBox1.Height - dispy Then
                drawRectangle.Y = PictureBox1.Height - dispy - drawRectangle.Height

            End If

            PictureBox1.Invalidate()
        End If
        dragxy_start.X = mousexy.X
        dragxy_start.Y = mousexy.Y

    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize


        If Not (currentImage Is Nothing) Then

            Dim dispwidth As Integer = PictureBox1.Width
            Dim dispheight As Integer = PictureBox1.Height
            Dim dispx As Integer
            Dim dispy As Integer



            If dispwidth / currentImage.Width < dispheight / currentImage.Height Then
                dispheight = dispwidth / currentImage.Width * currentImage.Height
                dispx = 0
                dispy = (PictureBox1.Height - dispheight) / 2
            Else

                dispwidth = dispheight / currentImage.Height * currentImage.Width
                dispx = (PictureBox1.Width - dispwidth) / 2
                dispy = 0

            End If

            zoomRatio = dispheight / currentImage.Height


            drawRectangle.Width = dispwidth
            drawRectangle.Height = dispheight
            drawRectangle.X = dispx
            drawRectangle.Y = dispy


            zoomRatio = dispwidth / currentImage.Width
            initzoomRatio = zoomRatio


            PictureBox1.Invalidate()
        End If

    End Sub

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        If drag_st = False Then


            drag_st = True


        End If



    End Sub

    Private Sub PictureBox1_MouseUp(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseUp
        drag_st = False
    End Sub

    Private Sub コピーToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles コピーToolStripMenuItem.Click
        Clipboard.SetImage(currentImage)
    End Sub



    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        Call frw_key()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Call rev_key()
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        If (_bScreenMode = False) Then
            Call Change_fullSCR()
        End If
    End Sub

    Private Sub PictureBox3_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox3.MouseEnter
        PictureBox3.Image = My.Resources.Resources.左アクティブ
    End Sub

    Private Sub PictureBox3_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox3.MouseLeave
        PictureBox3.Image = My.Resources.Resources.左非アクティブ
    End Sub

    Private Sub PictureBox5_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox5.MouseEnter
        PictureBox5.Image = My.Resources.Resources.右アクティブ
    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox5.MouseLeave
        PictureBox5.Image = My.Resources.Resources.右非アクティブ
    End Sub

    Private Sub PictureBox4_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox4.MouseEnter
        PictureBox4.Image = My.Resources.Resources.中央アクティブ

    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave
        PictureBox4.Image = My.Resources.Resources.中央非アクティブ

    End Sub


    Private Sub Form1_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles MyBase.PreviewKeyDown
        Console.WriteLine("フォームのキーイベント")
        Select Case e.KeyCode
            Case Keys.Escape
                e.IsInputKey = True

                If (Me.WindowState = FormWindowState.Maximized) Then
                    Me.WindowState = FormWindowState.Normal
                End If
                If (_bScreenMode = True) Then
                    Call Change_fullSCR()
                End If

            Case Keys.Left
                e.IsInputKey = True
                Console.WriteLine("左")
                Call rev_key()

            Case Keys.Right
                Console.WriteLine("右")

                e.IsInputKey = True
                Call frw_key()
            Case Keys.Enter
                e.IsInputKey = True
                Call Change_fullSCR()
        End Select
    End Sub

    Private Sub TableLayoutPanel1_MouseClick(sender As Object, e As MouseEventArgs) Handles TableLayoutPanel1.MouseClick
        Me.ActiveControl = Nothing
    End Sub

    Private Sub PictureBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles PictureBox2.MouseClick
        Me.ActiveControl = Nothing
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        AxWindowsMediaPlayer1.Ctlcontrols.stop()
    End Sub
End Class
