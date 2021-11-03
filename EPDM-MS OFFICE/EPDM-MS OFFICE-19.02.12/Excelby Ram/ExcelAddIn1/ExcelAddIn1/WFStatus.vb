
Imports EdmLib
Imports System.Drawing
Imports System.Windows.Forms

Public Class WFStatus
    Dim fname As String
    Dim xl As Excel.Application
    Dim wb As Excel.Workbook
    Dim vault As EdmVault5
    Dim vault1 As IEdmVault8
    Dim vinfo() As EdmViewInfo
    Dim fol As IEdmFolder5
    Dim dfile As IEdmFile5
    Dim iuser As IEdmUser5
    Dim wfs As IEdmState5
    Dim lst As IEdmBatchListing
    Dim lfile() As EdmListFile
    Dim cols() As EdmListCol
    Dim ia As Integer
    Dim vname As String
    Dim z As String
    
    Dim wrkmgr As IEdmWorkflowMgr6
    Dim wrkflw As IEdmWorkflow5
    Dim stat As IEdmState5
    Dim tran As IEdmTransition5
    Dim pos As IEdmPos5
    Dim fstat As String()
    Dim a As Integer = 1
    Dim tstat As String()
    Dim b As Integer = 1
    Dim lbl As Label() = New Label(11) {}
    Dim c As Integer = 1
    Dim x As Integer = 280
    Dim y As Integer = 25
    Dim i As Integer = 0
    Dim j As Integer
    Dim k As Integer
    Dim nstat As String



    Private Sub WFStatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        xl = GetObject(, "Excel.Application")
        wb = xl.ActiveWorkbook

        fname = wb.FullName
        vault = New EdmVault5
        vault1 = vault
        vault1.GetVaultViews(vinfo, False)

        For Me.ia = 0 To UBound(vinfo)
            If fname.Contains(vinfo(ia).mbsPath) = True Then
                vname = vinfo(ia).mbsVaultName
                Exit For
            End If
        Next

        vault.LoginAuto(vname, 0)
        fol = vault.GetFolderFromPath(wb.Path)
        dfile = vault.GetFileFromPath(fname, fol)
        lst = vault.CreateUtility(EdmUtility.EdmUtil_BatchList)
        lst.AddFile(fname, FileSystem.FileDateTime(fname), 0)
        lst.CreateList("", cols)
        lst.GetFiles(lfile)
        z = lfile(0).moCurrentState.mbsWorkflowName
        Dim rec As Rectangle() = New Rectangle(11) {}
        Dim y1 As Single = 20
        Dim x1 As Single = 250
        Dim w1 As Single = 250
        Dim h1 As Single = 30
        Dim d As Integer = 1
        Dim newfont As Font
        newfont = New Font("Sans Serif", 7, FontStyle.Bold)
        Dim points As PointF()
        Dim w As Integer = PictureBox1.Width
        Dim h As Integer = PictureBox1.Height
        Dim pen As New System.Drawing.Pen(System.Drawing.Color.Firebrick, 2)
        Dim pen2 As New System.Drawing.Pen(System.Drawing.Color.Black, 3)
        Dim pen1 As New System.Drawing.Pen(System.Drawing.Color.Black, 3)
        Dim bmp As New Bitmap(w, h, System.Drawing.Imaging.PixelFormat.Format64bppArgb)
       
        Using g As Graphics = Graphics.FromImage(bmp)
            pen1.StartCap = Drawing2D.LineCap.ArrowAnchor
            wrkmgr = vault.CreateUtility(EdmUtility.EdmUtil_WorkflowMgr)

            pos = wrkmgr.GetFirstWorkflowPosition
            While Not pos.IsNull
                wrkflw = wrkmgr.GetNextWorkflow(pos)
                If z = wrkflw.name Then
                    Exit While
                End If
            End While


            pos = wrkflw.GetFirstStatePosition
            While Not pos.IsNull
                stat = wrkflw.GetNextState(pos)
                Me.AutoScroll = True
                ReDim Preserve lbl(c)
                lbl(c) = New Label
                lbl(c).Visible = False
                lbl(c).Size = New System.Drawing.Size(150, 20)
                lbl(c).TextAlign = ContentAlignment.MiddleCenter
                lbl(c).ForeColor = System.Drawing.Color.Blue
                lbl(c).BackColor = System.Drawing.Color.Transparent
                PictureBox1.SendToBack()
                lbl(c).BringToFront()
                Me.Controls.Add(lbl(c))
                c += 1
                ReDim Preserve rec(d)
                rec(d) = New Rectangle(x1, y1, w1, h1)
                y1 = y1 + 80
                d += 1
                g.DrawRectangles(pen, rec)
            End While
            stat = wrkflw.GetState(wrkflw.InitialState.Name)
            pos = stat.GetFirstTransitionPosition
            While Not pos.IsNull
                tran = stat.GetNextTransition(pos)
                ReDim Preserve fstat(a)
                ReDim Preserve tstat(b)
                fstat(a) = tran.FromState.Name
                tstat(b) = tran.ToState.Name
                b += 1
                i += 1
            End While

            Select Case i
                Case 1
                    If tstat(1) <> "" Then
                        lbl(1).Text = fstat(1)
                        lbl(1).Location = New System.Drawing.Point(x, y)
                        lbl(1).Visible = True
                        PictureBox1.SendToBack()
                        lbl(1).BringToFront()
                        y = y + 80
                        lbl(2).Text = tstat(1)
                        lbl(2).Location = New System.Drawing.Point(x, y)
                        lbl(2).Visible = True
                        PictureBox1.SendToBack()
                        lbl(2).BringToFront()
                        y = y + 80
                        nstat = lbl(2).Text
                        j = 1
                        k = 1
                        If lbl(1).Text <> "" And lbl(2).Text <> "" Then
                            g.DrawLine(pen2, 375, 50, 375, 100)
                            g.DrawLine(pen1, 375, 100, 375, 90)
                        End If

                    End If

                Case 2
                    stat = wrkflw.GetState(tstat(1))
                    pos = stat.GetFirstTransitionPosition
                    While Not pos.IsNull
                        tran = stat.GetNextTransition(pos)
                        ReDim Preserve tstat(b)
                        tstat(b) = tran.ToState.Name
                        If tstat(b) = tstat(2) Then
                            lbl(1).Text = fstat(1)
                            lbl(1).Location = New System.Drawing.Point(x, y)
                            lbl(1).Visible = True
                            PictureBox1.SendToBack()
                            lbl(1).BringToFront()
                            y = y + 80
                            lbl(2).Text = tstat(1)
                            lbl(2).Location = New System.Drawing.Point(x, y)
                            lbl(2).Visible = True
                            PictureBox1.SendToBack()
                            lbl(2).BringToFront()
                            y = y + 80

                            lbl(3).Text = tstat(2)
                            lbl(3).Location = New System.Drawing.Point(x, y)
                            lbl(3).Visible = True
                            PictureBox1.SendToBack()
                            lbl(3).BringToFront()
                            y = y + 80
                            nstat = lbl(3).Text
                            j = 2
                            k = 2
                            If lbl(1).Text <> "" And lbl(2).Text <> "" And lbl(3).Text <> "" Then
                                g.DrawLine(pen2, 375, 50, 375, 100)
                                g.DrawLine(pen2, 375, 130, 375, 180)
                                g.DrawLine(pen1, 375, 100, 375, 90)
                                g.DrawLine(pen1, 375, 180, 375, 170)
                            End If

                        End If
                        b += 1

                    End While
                    stat = wrkflw.GetState(tstat(1))
                    pos = stat.GetFirstTransitionPosition
                    While Not pos.IsNull
                        tran = stat.GetNextTransition(pos)
                        ReDim Preserve tstat(b)
                        tstat(b) = tran.ToState.Name
                        If lbl(3).Text <> "" Then
                            If tstat(b) <> lbl(3).Text And tstat(b) = lbl(1).Text Then
                                points = {New PointF(250, 35), New PointF(225, 35), New PointF(225, 115), New PointF(250, 115)}
                                g.DrawLines(pen2, points)
                                g.DrawLine(pen1, 250, 35, 240, 35)

                            End If
                        End If
                        b += 1
                    End While

                    stat = wrkflw.GetState(tstat(2))
                    pos = stat.GetFirstTransitionPosition
                    While Not pos.IsNull
                        tran = stat.GetNextTransition(pos)
                        ReDim Preserve tstat(b)
                        tstat(b) = tran.ToState.Name
                        If tstat(b) = tstat(1) Then
                            lbl(1).Text = fstat(1)
                            lbl(1).Location = New System.Drawing.Point(x, y)
                            lbl(1).Visible = True
                            PictureBox1.SendToBack()
                            lbl(1).BringToFront()
                            y = y + 80
                            lbl(2).Text = tstat(2)
                            lbl(2).Location = New System.Drawing.Point(x, y)
                            lbl(2).Visible = True
                            PictureBox1.SendToBack()
                            lbl(2).BringToFront()
                            y = y + 80

                            lbl(3).Text = tstat(1)
                            lbl(3).Location = New System.Drawing.Point(x, y)
                            lbl(3).Visible = True
                            PictureBox1.SendToBack()
                            lbl(3).BringToFront()
                            y = y + 80
                            nstat = lbl(3).Text

                            j = 2
                            k = 2
                            If lbl(1).Text <> "" And lbl(2).Text <> "" And lbl(3).Text <> "" Then
                                g.DrawLine(pen2, 375, 50, 375, 100)
                                g.DrawLine(pen2, 375, 130, 375, 180)
                                g.DrawLine(pen1, 375, 100, 375, 90)
                                g.DrawLine(pen1, 375, 180, 375, 170)
                            End If
                        End If
                        b += 1
                    End While

                    stat = wrkflw.GetState(wrkflw.InitialState.Name)
                    pos = stat.GetFirstTransitionPosition
                    While Not pos.IsNull
                        tran = stat.GetNextTransition(pos)
                        ReDim Preserve tstat(b)
                        tstat(b) = tran.ToState.Name
                        If lbl(2).Text <> "" Then
                            If tstat(b) <> lbl(2).Text And tstat(b) = lbl(3).Text Then
                                points = {New PointF(500, 40), New PointF(525, 40), New PointF(525, 190), New PointF(500, 190)}
                                g.DrawLine(pen1, 500, 190, 510, 190)
                                g.DrawLines(pen2, points)
                            End If
                        End If
                        b += 1
                    End While
                    stat = wrkflw.GetState(tstat(2))
                    pos = stat.GetFirstTransitionPosition
                    While Not pos.IsNull
                        tran = stat.GetNextTransition(pos)
                        ReDim Preserve tstat(b)
                        tstat(b) = tran.ToState.Name
                        If lbl(3).Text <> "" Then
                            If tstat(b) <> lbl(3).Text And tstat(b) = lbl(1).Text Then
                                points = {New PointF(250, 35), New PointF(225, 35), New PointF(225, 115), New PointF(250, 115)}
                                g.DrawLines(pen2, points)
                                g.DrawLine(pen1, 250, 35, 240, 35)
                            End If
                        End If
                        b += 1
                    End While
            End Select

            If nstat <> "" Then
                stat = wrkflw.GetState(nstat)

                pos = stat.GetFirstTransitionPosition
                i = 0
                b = 1
                While Not pos.IsNull
                    tran = stat.GetNextTransition(pos)
                    ReDim Preserve tstat(b)
                    tstat(b) = tran.ToState.Name
                    i += 1
                    b += 1
                End While

                Select Case i
                    Case 1
                        If j = 1 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text Then
                                lbl(3).Text = tstat(1)
                                lbl(3).Location = New System.Drawing.Point(x, y)
                                lbl(3).Visible = True
                                PictureBox1.SendToBack()
                                lbl(3).BringToFront()
                                y = y + 80
                                j = 3
                                k = 3
                                If lbl(2).Text <> "" And lbl(3).Text <> "" Then
                                    g.DrawLine(pen2, 375, 130, 375, 180)
                                    g.DrawLine(pen1, 375, 180, 375, 170)
                                End If
                                nstat = lbl(3).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(250, 35), New PointF(225, 35), New PointF(225, 115), New PointF(250, 115)}
                                g.DrawLines(pen2, points)
                                g.DrawLine(pen1, 250, 35, 240, 35)
                                nstat = ""
                            End If
                        End If

                        If j = 2 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text Then
                                lbl(4).Text = tstat(1)
                                lbl(4).Location = New System.Drawing.Point(x, y)
                                lbl(4).Visible = True
                                PictureBox1.SendToBack()
                                lbl(4).BringToFront()
                                y = y + 80
                                j = 4
                                k = 4
                                If lbl(3).Text <> "" And lbl(4).Text <> "" Then
                                    g.DrawLine(pen2, 375, 210, 375, 260)
                                    g.DrawLine(pen1, 375, 260, 375, 250)
                                End If
                                nstat = lbl(4).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(250, 120), New PointF(225, 120), New PointF(225, 190), New PointF(250, 190)}
                                g.DrawLine(pen1, 250, 120, 240, 120)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(250, 20), New PointF(200, 20), New PointF(200, 175), New PointF(300, 175), New PointF(300, 180)}
                                g.DrawLine(pen1, 250, 20, 240, 20)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If
                        End If

                    Case 2
                        If j = 1 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text Then
                                    lbl(3).Text = tstat(b)
                                    lbl(3).Location = New System.Drawing.Point(x, y)
                                    lbl(3).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(3).BringToFront()
                                    y = y + 80
                                    j = 5
                                    k = 3
                                    If lbl(2).Text <> "" And lbl(3).Text <> "" Then
                                        g.DrawLine(pen2, 375, 130, 375, 180)
                                        g.DrawLine(pen1, 375, 180, 375, 170)
                                    End If
                                    nstat = lbl(3).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(250, 35), New PointF(225, 35), New PointF(225, 115), New PointF(250, 115)}
                                    g.DrawLines(pen2, points)
                                    g.DrawLine(pen1, 250, 35, 240, 35)
                                End If
                            End While
                        End If


                        If j = 2 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> "" And tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text Then
                                    lbl(4).Text = tstat(b)
                                    lbl(4).Location = New System.Drawing.Point(x, y)
                                    lbl(4).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(4).BringToFront()
                                    y = y + 80
                                    j = 6
                                    k = 4
                                    If lbl(3).Text <> "" And lbl(4).Text <> "" Then
                                        g.DrawLine(pen2, 375, 210, 375, 260)
                                        g.DrawLine(pen1, 375, 260, 375, 250)
                                    End If
                                    nstat = lbl(4).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(250, 120), New PointF(225, 120), New PointF(225, 190), New PointF(250, 190)}
                                    g.DrawLine(pen1, 250, 120, 240, 120)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(250, 20), New PointF(200, 20), New PointF(200, 175), New PointF(300, 175), New PointF(300, 180)}
                                    g.DrawLine(pen1, 250, 20, 240, 20)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                End If
                            End While
                        End If
                End Select

            End If



            If nstat <> "" Then
                stat = wrkflw.GetState(nstat)

                pos = stat.GetFirstTransitionPosition
                i = 0
                b = 1
                While Not pos.IsNull
                    tran = stat.GetNextTransition(pos)
                    ReDim Preserve tstat(b)
                    tstat(b) = tran.ToState.Name
                    i += 1
                    b += 1
                End While


                Select Case i

                    Case 1
                        If j = 3 Or j = 5 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text Then
                                lbl(4).Text = tstat(1)
                                lbl(4).Location = New System.Drawing.Point(x, y)
                                lbl(4).Visible = True
                                PictureBox1.SendToBack()
                                lbl(4).BringToFront()
                                y = y + 80
                                j = 7
                                k = 5
                                If lbl(3).Text <> "" And lbl(4).Text <> "" Then
                                    g.DrawLine(pen2, 375, 210, 375, 260)
                                    g.DrawLine(pen1, 375, 260, 375, 250)
                                End If
                                nstat = lbl(4).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(250, 20), New PointF(200, 20), New PointF(200, 175), New PointF(300, 175), New PointF(300, 180)}
                                g.DrawLine(pen1, 250, 20, 240, 20)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(250, 120), New PointF(225, 120), New PointF(225, 190), New PointF(250, 190)}
                                g.DrawLine(pen1, 250, 120, 240, 120)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If

                        End If


                        If j = 4 Or j = 6 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text Then
                                lbl(5).Text = tstat(1)
                                lbl(5).Location = New System.Drawing.Point(x, y)
                                lbl(5).Visible = True
                                PictureBox1.SendToBack()
                                lbl(5).BringToFront()
                                y = y + 80
                                j = 8
                                k = 6
                                If lbl(4).Text <> "" And lbl(5).Text <> "" Then
                                    g.DrawLine(pen2, 375, 290, 375, 340)
                                    g.DrawLine(pen1, 375, 340, 375, 330)
                                End If
                                nstat = lbl(5).Text

                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(300, 260), New PointF(300, 230), New PointF(180, 230), New PointF(180, 75), New PointF(300, 75), New PointF(300, 50)}
                                g.DrawLine(pen1, 300, 50, 300, 55)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(400, 260), New PointF(400, 230), New PointF(520, 230), New PointF(520, 140), New PointF(400, 140), New PointF(400, 130)}
                                g.DrawLine(pen1, 400, 130, 400, 135)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(250, 200), New PointF(225, 200), New PointF(225, 270), New PointF(250, 270)}
                                g.DrawLine(pen1, 250, 200, 240, 200)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If
                        End If


                    Case 2

                        If j = 3 Or j = 5 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text Then
                                    lbl(4).Text = tstat(b)
                                    lbl(4).Location = New System.Drawing.Point(x, y)
                                    lbl(4).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(4).BringToFront()
                                    y = y + 80
                                    j = 9
                                    k = 5
                                    If lbl(3).Text <> "" And lbl(4).Text <> "" Then
                                        g.DrawLine(pen2, 375, 210, 375, 260)
                                        g.DrawLine(pen1, 375, 260, 375, 250)
                                    End If
                                    nstat = lbl(4).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(250, 20), New PointF(200, 20), New PointF(200, 175), New PointF(300, 175), New PointF(300, 180)}
                                    g.DrawLine(pen1, 250, 20, 240, 20)
                                    g.DrawLines(pen2, points)

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(250, 120), New PointF(225, 120), New PointF(225, 190), New PointF(250, 190)}
                                    g.DrawLine(pen1, 250, 120, 240, 120)
                                    g.DrawLines(pen2, points)
                                End If
                            End While
                        End If


                        If j = 4 Or j = 6 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text Then
                                    lbl(5).Text = tstat(b)
                                    lbl(5).Location = New System.Drawing.Point(x, y)
                                    lbl(5).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(5).BringToFront()
                                    y = y + 80
                                    j = 10
                                    k = 6
                                    If lbl(4).Text <> "" And lbl(5).Text <> "" Then
                                        g.DrawLine(pen2, 375, 290, 375, 340)
                                        g.DrawLine(pen1, 375, 340, 375, 330)
                                    End If
                                    nstat = lbl(5).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(300, 260), New PointF(300, 230), New PointF(180, 230), New PointF(180, 75), New PointF(300, 75), New PointF(300, 50)}
                                    g.DrawLine(pen1, 300, 50, 300, 55)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(400, 260), New PointF(400, 230), New PointF(520, 230), New PointF(520, 140), New PointF(400, 140), New PointF(400, 130)}
                                    g.DrawLine(pen1, 400, 130, 400, 135)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(250, 200), New PointF(225, 200), New PointF(225, 270), New PointF(250, 270)}
                                    g.DrawLine(pen1, 250, 200, 240, 200)
                                    g.DrawLines(pen2, points)
                                End If
                            End While
                        End If
                End Select
            End If


            If nstat <> "" Then
                stat = wrkflw.GetState(nstat)

                pos = stat.GetFirstTransitionPosition
                i = 0
                b = 1
                While Not pos.IsNull
                    tran = stat.GetNextTransition(pos)
                    ReDim Preserve tstat(b)
                    tstat(b) = tran.ToState.Name
                    i += 1
                    b += 1
                End While


                Select Case i

                    Case 1
                        If j = 7 Or j = 9 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text Then
                                lbl(5).Text = tstat(1)
                                lbl(5).Location = New System.Drawing.Point(x, y)
                                lbl(5).Visible = True
                                PictureBox1.SendToBack()
                                lbl(5).BringToFront()
                                y = y + 80
                                j = 11
                                k = 7
                                If lbl(4).Text <> "" And lbl(5).Text <> "" Then
                                    g.DrawLine(pen2, 375, 290, 375, 340)
                                    g.DrawLine(pen1, 375, 340, 375, 330)
                                End If
                                nstat = lbl(5).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(300, 260), New PointF(300, 230), New PointF(180, 230), New PointF(180, 75), New PointF(300, 75), New PointF(300, 50)}
                                g.DrawLine(pen1, 300, 50, 300, 55)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(400, 260), New PointF(400, 230), New PointF(520, 230), New PointF(520, 140), New PointF(400, 140), New PointF(400, 130)}
                                g.DrawLine(pen1, 400, 130, 400, 135)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(250, 200), New PointF(225, 200), New PointF(225, 270), New PointF(250, 270)}
                                g.DrawLine(pen1, 250, 200, 240, 200)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If
                        End If


                        If j = 8 Or j = 10 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text And tstat(1) <> lbl(4).Text Then
                                lbl(6).Text = tstat(1)
                                lbl(6).Location = New System.Drawing.Point(x, y)
                                lbl(6).Visible = True
                                PictureBox1.SendToBack()
                                lbl(6).BringToFront()
                                y = y + 80
                                j = 12
                                k = 8
                                If lbl(5).Text <> "" And lbl(6).Text <> "" Then
                                    g.DrawLine(pen2, 375, 370, 375, 420)
                                    g.DrawLine(pen1, 375, 420, 375, 410)
                                End If
                                nstat = lbl(6).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(250, 40), New PointF(190, 40), New PointF(190, 355), New PointF(250, 355)}
                                g.DrawLine(pen1, 250, 40, 240, 40)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(500, 115), New PointF(580, 115), New PointF(580, 355), New PointF(500, 355)}
                                g.DrawLine(pen1, 500, 115, 510, 115)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(475, 340), New PointF(475, 320), New PointF(560, 320), New PointF(560, 240), New PointF(475, 240), New PointF(475, 210)}
                                g.DrawLine(pen1, 475, 210, 475, 220)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(4).Text Then
                                points = {New PointF(250, 280), New PointF(225, 280), New PointF(225, 350), New PointF(250, 350)}
                                g.DrawLine(pen1, 250, 280, 240, 280)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If
                        End If


                    Case 2

                        If j = 7 Or j = 9 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text Then
                                    lbl(5).Text = tstat(b)
                                    lbl(5).Location = New System.Drawing.Point(x, y)
                                    lbl(5).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(5).BringToFront()
                                    y = y + 80
                                    j = 13
                                    k = 7
                                    If lbl(4).Text <> "" And lbl(5).Text <> "" Then
                                        g.DrawLine(pen2, 375, 290, 375, 340)
                                        g.DrawLine(pen1, 375, 340, 375, 330)
                                    End If
                                    nstat = lbl(5).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(300, 260), New PointF(300, 230), New PointF(180, 230), New PointF(180, 75), New PointF(300, 75), New PointF(300, 50)}
                                    g.DrawLine(pen1, 300, 50, 300, 55)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(400, 260), New PointF(400, 230), New PointF(520, 230), New PointF(520, 140), New PointF(400, 140), New PointF(400, 130)}
                                    g.DrawLine(pen1, 400, 130, 400, 135)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(250, 200), New PointF(225, 200), New PointF(225, 270), New PointF(250, 270)}
                                    g.DrawLine(pen1, 250, 200, 240, 200)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                End If
                            End While

                        End If



                        If j = 8 Or j = 10 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text And tstat(b) <> lbl(4).Text Then
                                    lbl(6).Text = tstat(b)
                                    lbl(6).Location = New System.Drawing.Point(x, y)
                                    lbl(6).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(6).BringToFront()
                                    y = y + 80
                                    j = 14
                                    k = 8
                                    If lbl(5).Text <> "" And lbl(6).Text <> "" Then
                                        g.DrawLine(pen2, 375, 370, 375, 420)
                                        g.DrawLine(pen1, 375, 420, 375, 410)
                                    End If
                                    nstat = lbl(6).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(250, 40), New PointF(190, 40), New PointF(190, 355), New PointF(250, 355)}
                                    g.DrawLine(pen1, 250, 40, 240, 40)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(500, 115), New PointF(580, 115), New PointF(580, 355), New PointF(500, 355)}
                                    g.DrawLine(pen1, 500, 115, 510, 115)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(475, 340), New PointF(475, 320), New PointF(560, 320), New PointF(560, 240), New PointF(475, 240), New PointF(475, 210)}
                                    g.DrawLine(pen1, 475, 210, 475, 220)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(4).Text Then
                                    points = {New PointF(250, 280), New PointF(225, 280), New PointF(225, 350), New PointF(250, 350)}
                                    g.DrawLine(pen1, 250, 280, 240, 280)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                End If
                            End While
                        End If
                End Select
            End If




            If nstat <> "" Then
                stat = wrkflw.GetState(nstat)

                pos = stat.GetFirstTransitionPosition
                i = 0
                b = 1
                While Not pos.IsNull
                    tran = stat.GetNextTransition(pos)
                    ReDim Preserve tstat(b)
                    tstat(b) = tran.ToState.Name
                    i += 1
                    b += 1
                End While


                Select Case i

                    Case 1
                        If j = 11 Or j = 13 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text And tstat(1) <> lbl(4).Text Then
                                lbl(6).Text = tstat(1)
                                lbl(6).Location = New System.Drawing.Point(x, y)
                                lbl(6).Visible = True
                                PictureBox1.SendToBack()
                                lbl(6).BringToFront()
                                y = y + 80
                                j = 15
                                k = 9
                                If lbl(5).Text <> "" And lbl(6).Text <> "" Then
                                    g.DrawLine(pen2, 375, 370, 375, 420)
                                    g.DrawLine(pen1, 375, 420, 375, 410)
                                End If
                                nstat = lbl(6).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(250, 40), New PointF(190, 40), New PointF(190, 355), New PointF(250, 355)}
                                g.DrawLine(pen1, 250, 40, 240, 40)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(500, 115), New PointF(580, 115), New PointF(580, 355), New PointF(500, 355)}
                                g.DrawLine(pen1, 500, 115, 510, 115)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(475, 340), New PointF(475, 320), New PointF(560, 320), New PointF(560, 240), New PointF(475, 240), New PointF(475, 210)}
                                g.DrawLine(pen1, 475, 210, 475, 220)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(4).Text Then
                                points = {New PointF(250, 280), New PointF(225, 280), New PointF(225, 350), New PointF(250, 350)}
                                g.DrawLine(pen1, 250, 280, 240, 280)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If
                        End If


                        If j = 12 Or j = 14 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text And tstat(1) <> lbl(4).Text And tstat(1) <> lbl(5).Text Then
                                lbl(7).Text = tstat(1)
                                lbl(7).Location = New System.Drawing.Point(x, y)
                                lbl(7).Visible = True
                                PictureBox1.SendToBack()
                                lbl(7).BringToFront()
                                y = y + 80
                                j = 16
                                k = 10
                                If lbl(6).Text <> "" And lbl(7).Text <> "" Then
                                    g.DrawLine(pen2, 375, 450, 375, 500)
                                    g.DrawLine(pen1, 375, 500, 375, 490)
                                End If
                                nstat = lbl(7).Text

                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(500, 435), New PointF(587, 435), New PointF(587, 37), New PointF(500, 37)}
                                g.DrawLine(pen1, 500, 37, 510, 37)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(250, 435), New PointF(177, 435), New PointF(177, 122), New PointF(250, 122)}
                                g.DrawLine(pen1, 250, 122, 240, 122)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(255, 420), New PointF(255, 400), New PointF(227, 400), New PointF(227, 256), New PointF(255, 256), New PointF(255, 210)}
                                g.DrawLine(pen1, 255, 210, 255, 215)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(4).Text Then
                                points = {New PointF(460, 420), New PointF(460, 400), New PointF(537, 400), New PointF(537, 307), New PointF(460, 307), New PointF(460, 290)}
                                g.DrawLine(pen1, 460, 290, 460, 295)
                                g.DrawLines(pen2, points)
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(5).Text Then
                                points = {New PointF(250, 360), New PointF(225, 360), New PointF(225, 430), New PointF(250, 430)}
                                g.DrawLine(pen1, 250, 360, 240, 360)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If

                        End If


                    Case 2

                        If j = 11 Or j = 13 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text And tstat(b) <> lbl(4).Text Then
                                    lbl(6).Text = tstat(b)
                                    lbl(6).Location = New System.Drawing.Point(x, y)
                                    lbl(6).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(6).BringToFront()
                                    y = y + 80
                                    j = 17
                                    k = 9
                                    If lbl(5).Text <> "" And lbl(6).Text <> "" Then
                                        g.DrawLine(pen2, 375, 370, 375, 420)
                                        g.DrawLine(pen1, 375, 420, 375, 410)
                                    End If
                                    nstat = lbl(6).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(250, 40), New PointF(190, 40), New PointF(190, 355), New PointF(250, 355)}
                                    g.DrawLine(pen1, 250, 40, 240, 40)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(500, 115), New PointF(580, 115), New PointF(580, 355), New PointF(500, 355)}
                                    g.DrawLine(pen1, 500, 115, 510, 115)
                                    nstat = ""
                                    g.DrawLines(pen2, points)

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(475, 340), New PointF(475, 320), New PointF(560, 320), New PointF(560, 240), New PointF(475, 240), New PointF(475, 210)}
                                    g.DrawLine(pen1, 475, 210, 475, 220)
                                    g.DrawLines(pen2, points)
                                    nstat = ""

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(4).Text Then
                                    points = {New PointF(250, 280), New PointF(225, 280), New PointF(225, 350), New PointF(250, 350)}
                                    g.DrawLine(pen1, 250, 280, 240, 280)
                                    g.DrawLines(pen2, points)
                                    nstat = ""
                                End If
                            End While

                        End If


                        If j = 8 Or j = 10 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text And tstat(b) <> lbl(4).Text And tstat(b) <> lbl(5).Text Then
                                    lbl(7).Text = tstat(b)
                                    lbl(7).Location = New System.Drawing.Point(x, y)
                                    lbl(7).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(7).BringToFront()
                                    y = y + 80
                                    j = 18
                                    k = 10
                                    If lbl(6).Text <> "" And lbl(7).Text <> "" Then
                                        g.DrawLine(pen2, 375, 450, 375, 500)
                                        g.DrawLine(pen1, 375, 500, 375, 490)
                                    End If
                                    nstat = lbl(7).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(500, 435), New PointF(587, 435), New PointF(587, 37), New PointF(500, 37)}
                                    g.DrawLine(pen1, 500, 37, 510, 37)
                                    g.DrawLines(pen2, points)

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(250, 435), New PointF(177, 435), New PointF(177, 122), New PointF(250, 122)}
                                    g.DrawLine(pen1, 250, 122, 240, 122)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(255, 420), New PointF(255, 400), New PointF(227, 400), New PointF(227, 256), New PointF(255, 256), New PointF(255, 210)}
                                    g.DrawLine(pen1, 255, 210, 255, 215)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(4).Text Then
                                    points = {New PointF(460, 420), New PointF(460, 400), New PointF(537, 400), New PointF(537, 307), New PointF(460, 307), New PointF(460, 290)}
                                    g.DrawLine(pen1, 460, 290, 460, 295)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(5).Text Then
                                    points = {New PointF(250, 360), New PointF(225, 360), New PointF(225, 430), New PointF(250, 430)}
                                    g.DrawLine(pen1, 250, 360, 240, 360)
                                    g.DrawLines(pen2, points)

                                End If
                            End While
                        End If
                End Select
            End If



            If nstat <> "" Then
                stat = wrkflw.GetState(nstat)

                pos = stat.GetFirstTransitionPosition
                i = 0
                b = 1
                While Not pos.IsNull
                    tran = stat.GetNextTransition(pos)
                    ReDim Preserve tstat(b)
                    tstat(b) = tran.ToState.Name
                    i += 1
                    b += 1
                End While


                Select Case i

                    Case 1
                        If j = 15 Or j = 17 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text And tstat(1) <> lbl(4).Text And tstat(1) <> lbl(5).Text And tstat(1) <> lbl(6).Text Then
                                lbl(7).Text = tstat(1)
                                lbl(7).Location = New System.Drawing.Point(x, y)
                                lbl(7).Visible = True
                                PictureBox1.SendToBack()
                                lbl(7).BringToFront()
                                y = y + 80
                                j = 19
                                k = 11
                                If lbl(6).Text <> "" And lbl(7).Text <> "" Then
                                    g.DrawLine(pen2, 375, 450, 375, 500)
                                    g.DrawLine(pen1, 375, 500, 375, 490)
                                End If
                                nstat = lbl(7).Text
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(500, 435), New PointF(587, 435), New PointF(587, 37), New PointF(500, 37)}
                                g.DrawLine(pen1, 500, 37, 510, 37)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(250, 435), New PointF(177, 435), New PointF(177, 122), New PointF(250, 122)}
                                g.DrawLine(pen1, 250, 122, 240, 122)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(255, 420), New PointF(255, 400), New PointF(227, 400), New PointF(227, 256), New PointF(255, 256), New PointF(255, 210)}
                                g.DrawLine(pen1, 255, 210, 255, 215)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(4).Text Then
                                points = {New PointF(460, 420), New PointF(460, 400), New PointF(537, 400), New PointF(537, 307), New PointF(460, 307), New PointF(460, 290)}
                                g.DrawLine(pen1, 460, 290, 460, 295)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(5).Text Then
                                points = {New PointF(250, 360), New PointF(225, 360), New PointF(225, 430), New PointF(250, 430)}
                                g.DrawLine(pen1, 250, 360, 240, 360)
                                g.DrawLines(pen2, points)
                                nstat = ""
                            End If

                        End If


                        If j = 16 Or j = 18 Then
                            If tstat(1) <> "" And tstat(1) <> lbl(1).Text And tstat(1) <> lbl(2).Text And tstat(1) <> lbl(3).Text And tstat(1) <> lbl(4).Text And tstat(1) <> lbl(5).Text And tstat(1) <> lbl(6).Text And tstat(1) <> lbl(7).Text Then
                                lbl(8).Text = tstat(1)
                                lbl(8).Location = New System.Drawing.Point(x, y)
                                lbl(8).Visible = True
                                PictureBox1.SendToBack()
                                lbl(8).BringToFront()
                                y = y + 80
                                j = 20
                                k = 12
                                If lbl(7).Text <> "" And lbl(8).Text <> "" Then
                                    g.DrawLine(pen2, 375, 530, 375, 580)
                                    g.DrawLine(pen1, 375, 580, 375, 570)
                                End If
                                nstat = lbl(8).Text

                            ElseIf tstat(1) <> "" And tstat(1) = lbl(1).Text Then
                                points = {New PointF(250, 515), New PointF(203, 515), New PointF(203, 44), New PointF(250, 44)}
                                g.DrawLine(pen1, 250, 44, 240, 44)
                                g.DrawLines(Pens.Green, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(2).Text Then
                                points = {New PointF(500, 515), New PointF(515, 515), New PointF(515, 117), New PointF(500, 117)}
                                g.DrawLine(pen1, 500, 117, 510, 117)
                                g.DrawLines(Pens.Green, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(3).Text Then
                                points = {New PointF(285, 500), New PointF(285, 480), New PointF(167, 480), New PointF(167, 239), New PointF(285, 239), New PointF(285, 210)}
                                g.DrawLine(pen1, 285, 210, 285, 215)
                                g.DrawLines(Pens.Green, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(4).Text Then

                                points = {New PointF(250, 505), New PointF(152, 505), New PointF(152, 333), New PointF(267, 333), New PointF(267, 290)}
                                g.DrawLine(pen1, 267, 290, 267, 295)
                                g.DrawLines(Pens.Green, points)
                                nstat = ""

                            ElseIf tstat(1) <> "" And tstat(1) = lbl(5).Text Then
                                points = {New PointF(460, 500), New PointF(460, 480), New PointF(567, 480), New PointF(567, 383), New PointF(460, 383), New PointF(460, 370)}
                                g.DrawLine(pen1, 460, 370, 460, 375)
                                g.DrawLines(Pens.Green, points)
                                nstat = ""
                            ElseIf tstat(1) <> "" And tstat(1) = lbl(6).Text Then
                                points = {New PointF(250, 440), New PointF(225, 440), New PointF(225, 510), New PointF(250, 510)}
                                g.DrawLine(pen1, 250, 440, 240, 440)
                                g.DrawLines(Pens.Blue, points)
                                nstat = ""
                            End If
                        End If


                    Case 2

                        If j = 15 Or j = 17 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text And tstat(b) <> lbl(4).Text And tstat(b) <> lbl(5).Text Then
                                    lbl(7).Text = tstat(b)
                                    lbl(7).Location = New System.Drawing.Point(x, y)
                                    lbl(7).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(7).BringToFront()
                                    y = y + 80
                                    j = 21
                                    k = 11
                                    If lbl(6).Text <> "" And lbl(7).Text <> "" Then
                                        g.DrawLine(pen2, 375, 450, 375, 500)
                                        g.DrawLine(pen1, 375, 500, 375, 490)
                                    End If
                                    nstat = lbl(7).Text

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(500, 435), New PointF(587, 435), New PointF(587, 37), New PointF(500, 37)}
                                    g.DrawLine(pen1, 500, 37, 510, 37)
                                    g.DrawLines(pen2, points)

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(250, 435), New PointF(177, 435), New PointF(177, 122), New PointF(250, 122)}
                                    g.DrawLine(pen1, 250, 122, 240, 122)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(255, 420), New PointF(255, 400), New PointF(227, 400), New PointF(227, 256), New PointF(255, 256), New PointF(255, 210)}
                                    g.DrawLine(pen1, 255, 210, 255, 215)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(4).Text Then
                                    points = {New PointF(460, 420), New PointF(460, 400), New PointF(537, 400), New PointF(537, 307), New PointF(460, 307), New PointF(460, 290)}
                                    g.DrawLine(pen1, 460, 290, 460, 295)
                                    g.DrawLines(pen2, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(5).Text Then
                                    points = {New PointF(250, 360), New PointF(225, 360), New PointF(225, 430), New PointF(250, 430)}
                                    g.DrawLine(pen1, 250, 360, 240, 360)
                                    g.DrawLines(pen2, points)
                                End If
                            End While
                        End If


                        If j = 16 Or j = 18 Then
                            stat = wrkflw.GetState(nstat)
                            nstat = ""
                            pos = stat.GetFirstTransitionPosition
                            b = 1
                            While Not pos.IsNull
                                tran = stat.GetNextTransition(pos)
                                ReDim Preserve tstat(b)
                                tstat(b) = tran.ToState.Name
                                If tstat(b) <> lbl(1).Text And tstat(b) <> lbl(2).Text And tstat(b) <> lbl(3).Text And tstat(b) <> lbl(4).Text And tstat(b) <> lbl(5).Text And tstat(b) <> lbl(6).Text Then
                                    lbl(8).Text = tstat(b)
                                    lbl(8).Location = New System.Drawing.Point(x, y)
                                    lbl(8).Visible = True
                                    PictureBox1.SendToBack()
                                    lbl(8).BringToFront()
                                    y = y + 80
                                    j = 22
                                    k = 12
                                    If lbl(7).Text <> "" And lbl(8).Text <> "" Then
                                        g.DrawLine(pen2, 375, 530, 375, 580)
                                        g.DrawLine(pen1, 375, 580, 375, 570)
                                    End If
                                    nstat = lbl(8).Text
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(1).Text Then
                                    points = {New PointF(250, 515), New PointF(203, 515), New PointF(203, 44), New PointF(250, 44)}
                                    g.DrawLine(pen1, 250, 44, 240, 44)
                                    g.DrawLines(Pens.Green, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(2).Text Then
                                    points = {New PointF(500, 515), New PointF(515, 515), New PointF(515, 117), New PointF(500, 117)}
                                    g.DrawLine(pen1, 500, 117, 510, 117)
                                    g.DrawLines(Pens.Green, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(3).Text Then
                                    points = {New PointF(285, 500), New PointF(285, 480), New PointF(167, 480), New PointF(167, 239), New PointF(285, 239), New PointF(285, 210)}
                                    g.DrawLine(pen1, 285, 210, 285, 215)
                                    g.DrawLines(Pens.Green, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(4).Text Then

                                    points = {New PointF(250, 505), New PointF(152, 505), New PointF(152, 333), New PointF(267, 333), New PointF(267, 290)}
                                    g.DrawLine(pen1, 267, 290, 267, 295)
                                    g.DrawLines(Pens.Green, points)

                                ElseIf tstat(b) <> "" And tstat(b) = lbl(5).Text Then
                                    points = {New PointF(460, 500), New PointF(460, 480), New PointF(567, 480), New PointF(567, 383), New PointF(460, 383), New PointF(460, 370)}
                                    g.DrawLine(pen1, 460, 370, 460, 375)
                                    g.DrawLines(Pens.Green, points)
                                ElseIf tstat(b) <> "" And tstat(b) = lbl(6).Text Then
                                    points = {New PointF(250, 440), New PointF(225, 440), New PointF(225, 510), New PointF(250, 510)}
                                    g.DrawLine(pen1, 250, 440, 240, 440)
                                    g.DrawLines(Pens.Blue, points)

                                End If
                            End While
                        End If
                End Select
            End If
            PictureBox1.Image = bmp
        End Using
    End Sub



End Class