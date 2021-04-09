Imports System.Xml
Imports HtmlAgilityPack
Imports SequelMed.Core
Imports System.IO
Imports SequelMed.Core.Pattern
Imports TXTextControl
Imports System.Data.OleDb
Imports JUST
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Text.RegularExpressions
Imports System.Net
Imports SequelMed.Core.DB
Imports Oracle.ManagedDataAccess.Client
Imports SautinSoft.Document



Public Class Form1


    Public DBUserName As String = "CLOUDPEHR3"
    Public DBPassword As String = "MUGHAL"
    Public DBServer As String = "mu7550"



    Dim appCtx As Model.AppContext = Model.AppContext.Instance
    Dim dtSectionUnit As DataTable = New DataTable

    Private Const connstring As String = "Provider=Microsoft.Jet.OLEDB.4.0;" &
        "Data Source=D:\KaroVisits.xlsx;Extended Properties=""Excel 8.0;HDR=YES;"""
    Private Const connstringxslx As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=D:\KaroVisits1.xlsx;Extended Properties=""Excel 12.0;HDR=YES;"""
    Private Const sjfavisitString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=D:\SJFAVisit.xlsx;Extended Properties=""Excel 12.0;HDR=YES;"""
    Private Const ccpcvisitString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=D:\visit fixing\Cloud3Visit1.xlsx;Extended Properties=""Excel 12.0;HDR=YES;"""
    Private Const SWvisitString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=D:\SWVisitList.xlsx;Extended Properties=""Excel 12.0;HDR=YES;"""
    Private Const ExcelvisitString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=D:\Currently working\Excel Issue\ExcelCovidVisit.xlsx;Extended Properties=""Excel 12.0;HDR=YES;"""
    Private Const PSANvisits As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
        "Data Source=D:\Cloud3Visit1.xlsx;Extended Properties=""Excel 12.0;HDR=YES;"""

    Private Sub btnFileOpen_Click(sender As Object, e As EventArgs) Handles btnFileOpen.Click
        Dim fileBr As New OpenFileDialog
        fileBr.ShowDialog()
        If Not String.IsNullOrEmpty(fileBr.FileName) AndAlso IO.File.Exists(fileBr.FileName) Then
            txtFileName.Text = fileBr.FileName
        End If
    End Sub

    Private Sub btnChangeDOS_Click(sender As Object, e As EventArgs) Handles btnChangeDOS.Click
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = connstring
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visits$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows

            If ChangeDOS(rows) Then

            End If

        Next
        If True Then

        End If

    End Sub

    Private Function ChangeDOS(ByVal strRTFFilePath As DataRow) As Boolean
        Dim visitID As String = strRTFFilePath.Item("CLINICAL_VISIT_SEQ_NUM").ToString
        Using TmpTextcontrol As New TXTextControl.ServerTextControl
            If Not TmpTextcontrol.Create Then
                Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
            End If

            TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)

            For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields
                If tmpfield.Name.StartsWith("CLINICAL_VISIT_DATE") Then
                    If tmpfield.Text = Convert.ToDateTime(strRTFFilePath.Item("NEW_DOS").ToString).ToShortDateString Then
                        Exit Function
                    Else
                        tmpfield.Text = Convert.ToDateTime(strRTFFilePath.Item("NEW_DOS").ToString).ToShortDateString
                    End If

                End If
            Next



            Dim PageCount As Integer = 0
            Dim inputImages As New ArrayList()
            For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                Dim image As New MemoryStream()
                ' get the image from TX Text Control's page rendering engine
                Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                ' save and add the image to the ArrayList
                mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                inputImages.Add(image)
                PageCount += 1
            Next
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\" & visitID & "xmls.zip", strRTFFilePath.Item("Path") & "\" & visitID & "xmls-Bkp1.zip")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\ClinicalVisit.tif", strRTFFilePath.Item("Path") & "\ClinicalVisit-Bkp1.tif")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\ClinicalVisit.rtf", strRTFFilePath.Item("Path") & "\ClinicalVisit-Bkp1.rtf")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\ClinicalVisit.pdf", strRTFFilePath.Item("Path") & "\ClinicalVisit-Bkp1.pdf")
            'appCtx.DocSvr.EDM.FileSystem.File.Delete(strRTFFilePath & "\ClinicalVisit.pdf")
            Decorator.DoWithTempFile(Sub(tmp)
                                         CreateMultipageTIF(inputImages, tmp)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("Path") & "\ClinicalVisit.tif")
                                         TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("Path") & "\ClinicalVisit.rtf")
                                     End Sub)
        End Using
    End Function
    Private Sub CreateMultipageTIF(ByVal InputImages As ArrayList, ByVal Filename As String)
        ' set the image codec
        Dim info As Imaging.ImageCodecInfo = Nothing
        For Each ice As Imaging.ImageCodecInfo In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
            If ice.MimeType = "image/tiff" Then
                info = ice
                Exit For
            End If
        Next
        Dim ep As New Imaging.EncoderParameters(2)
        Dim firstPage As Boolean = True
        Dim img As System.Drawing.Image = Nothing
        'create a image instance from the 1st image
        For nLoopfile As Integer = 0 To InputImages.Count - 1
            'get image from src file
            Dim img_src As System.Drawing.Image = System.Drawing.Image.FromStream(DirectCast(InputImages(nLoopfile), Stream))
            Dim guid As Guid = img_src.FrameDimensionsList(0)
            Dim dimension As New System.Drawing.Imaging.FrameDimension(guid)
            'get the frames from src file
            For nLoopFrame As Integer = 0 To img_src.GetFrameCount(dimension) - 1
                img_src.SelectActiveFrame(dimension, nLoopFrame)
                ep.Param(0) = New Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Compression, Convert.ToInt32(System.Drawing.Imaging.EncoderValue.CompressionLZW))
                ' if first page, then create the initial image
                If firstPage Then
                    img = img_src

                    ep.Param(1) = New Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, Convert.ToInt32(System.Drawing.Imaging.EncoderValue.MultiFrame))
                    img.Save(Filename, info, ep)
                    firstPage = False
                    Continue For
                End If
                ' add image to the next frame
                ep.Param(1) = New Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, Convert.ToInt32(System.Drawing.Imaging.EncoderValue.FrameDimensionPage))
                img.SaveAdd(img_src, ep)
            Next
        Next
        ep.Param(1) = New Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, Convert.ToInt32(System.Drawing.Imaging.EncoderValue.Flush))
        img.SaveAdd(ep)
    End Sub

    Private Sub btnGetSectionData_Click(sender As Object, e As EventArgs) Handles btnGetSectionData.Click
        dtSectionUnit.Columns.Add("VISIT_SEQ_NUM")
        dtSectionUnit.Columns.Add("DESCRIPTION")
        dtSectionUnit.Columns.Add("QUESTION_SEQ_NUM")
        dtSectionUnit.Columns.Add("PICK_LIST_ID")
        dtSectionUnit.Columns.Add("PICK_LIST_NAME")
        dtSectionUnit.Columns.Add("ANSWER_VALUE_UNIT")
        Dim Path As String = "PEHR6VisitList.txt"
        Dim readText() As String = File.ReadAllLines(Path)
        For i As Integer = 0 To readText.Length - 1
            Application.DoEvents()
            Dim visitpath As String = readText(i)
            Label1.Text = "Processing " & (i + 1) & " of " & readText.Length
            If Not visitpath.Contains("Fixed") Then
                If GetSectionData(visitpath) Then
                    'My.Computer.FileSystem.WriteAllText(Path, My.Computer.FileSystem.ReadAllText(Path).Replace(visitpath, ""), False)
                    dtSectionUnit.WriteCSV("PEHR6VisitList.csv")
                End If
            End If
        Next
    End Sub

    Private Function GetSectionData(ByVal strRTFFilePath As String) As Boolean
        Dim visitID As String = Path.GetFileName(strRTFFilePath)
        Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(strRTFFilePath, "*.xml")
        If Not files.Empty() Then
            For Each fileName As String In files
                Try
                    Dim ds As New DataSet
                    If fileName.Contains("108948905166") OrElse fileName.Contains("10612335166") Then
                        Using ms As New MemoryStream(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(fileName))
                            ds.ReadXml(ms)
                        End Using
                        If ds.Tables.Contains("QUESTION_ROW") Then
                            For Each row As DataRow In ds.Tables("QUESTION_ROW").Select("DISPLAY_OBJECT = 'Procedure Control' and ANSWER_VALUE_UNIT is not null")
                                If row.Item("DISPLAY_OBJECT").ToString = "Procedure Control" AndAlso Not String.IsNullOrEmpty(row.Item("ANSWER_VALUE_UNIT").ToString) Then
                                    dtSectionUnit.Rows.Add(visitID, row.Item("DESCRIPTION").ToString, row.Item("QUESTION_SEQ_NUM").ToString, row.Item("PICK_LIST_ID").ToString, row.Item("PICK_LIST_NAME").ToString, row.Item("ANSWER_VALUE_UNIT").ToString)
                                End If
                            Next
                        End If
                    End If
                Catch ex As Exception
                End Try
            Next
        End If

        Return True
    End Function

    Private Sub BtnGetCurrptFiles_Click(sender As Object, e As EventArgs) Handles BtnGetCurrptFiles.Click
        GetCurrptFiles(txtFileName.Text)
    End Sub

    Private Function GetCurrptFiles(ByVal strRTFFilePath As String) As Boolean
        Dim folders As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetDirectories(strRTFFilePath)
        If Not folders.Empty() Then
            For Each fileName As String In folders
                Try

                Catch ex As Exception
                End Try
            Next
        End If
        Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(strRTFFilePath, "*.xml", [option]:=SearchOption.AllDirectories)
        If Not files.Empty() Then
            For Each fileName As String In files
                Try

                Catch ex As Exception
                End Try
            Next
        End If

        Return True
    End Function
    Dim ds As DataSet
    Private Sub btnGetUniqSec_Click(sender As Object, e As EventArgs) Handles btnGetUniqSec.Click
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter

        Dim Section As String = ""
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = connstringxslx
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows

            If String.IsNullOrEmpty(Section) Then
                Section = rows.Item("NOTE SECTION TYPE").ToString
            Else
                If Not Section.Contains(rows.Item("NOTE SECTION TYPE").ToString) Then
                    Section = Section + "," + rows.Item("NOTE SECTION TYPE").ToString
                End If
            End If

        Next

        Section = ""

        For Each rows As DataRow In ds.Tables(0).Rows
            'If rows.Item("NOTE STATUS").ToString.StartsWith("com") Then
            If String.IsNullOrEmpty(Section) Then
                Section = rows.Item("NOTE_ID").ToString
            Else
                If Not Section.Contains(rows.Item("NOTE_ID").ToString) Then
                    Section = Section + "," + rows.Item("NOTE_ID").ToString
                End If
            End If
            ' End If
        Next
        Dim code_des As String() = Section.Split(CChar(","))
        For Each visit As String In code_des
            Save(visit)
        Next
    End Sub
    Dim Index As Integer = 0
    Dim dt As DataTable = New DataTable()
    Private TmpTextControl As ServerTextControl = Nothing
    Dim folderPath As String = "D:\PatientHtml"
    Public Sub Save(strVisit As String)

        Dim dv As New DataView(ds.Tables(0))
        dv.RowFilter = "NOTE_ID = '" + strVisit + "'"
        dv.Sort = " DISPLAY_ORDER ASC "
        Dim dtVisit As DataTable = dv.ToTable
        Dim strMessage As String
        Dim imgpath As String = "http://localhost/CCDAView/test.jpg"
        Dim imgsrc As String = """data:Image/jpeg;base64,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"""
        strMessage = "<HTML><HEAD></HEAD><BODY><img src=""" + imgpath + """" + "alt=test.jpg"" style=" + "max-width:100%;" + ">"
        strMessage = strMessage + "<br/>"
        If dtVisit.Rows(0).Item("NOTE STATUS").ToString.StartsWith("com") Then
            strMessage = strMessage + "<center><b>" + dtVisit.Rows(0).Item("COMPANY NAME").ToString + "</b></center> "
        Else
            strMessage = strMessage + "<center><b><font size=" + "5" + " color=" + "red" + ">(Incomplete) </font></b><br/><b>" + dtVisit.Rows(0).Item("COMPANY NAME").ToString + "</b></center> "
        End If
        strMessage = strMessage + "<br/>"
        strMessage = strMessage + "<b>Patient's Name:</b>" + "     " + dtVisit.Rows(0).Item("LAST NAME").ToString + ", " + dtVisit.Rows(0).Item("FIRST NAME").ToString
        strMessage = strMessage + "<br/>"
        strMessage = strMessage + "<b>Date of Birth</b>" + "     " + dtVisit.Rows(0).Item("DATE OF BIRTH").ToString
        strMessage = strMessage + "<br/>"
        strMessage = strMessage + "<b>Clinical Visit Date:</b>" + "     " + dtVisit.Rows(0).Item("DATE OF VISIT").ToString
        strMessage = strMessage + "<br/><br/><br/>"
        Try
            For Each rows As DataRow In dtVisit.Rows
                If Not rows.Item("DATA").ToString.StartsWith("{") Then
                    rows.Item("DATA") = rows.Item("DATA").ToString.Replace("<b r>", "<br>")
                    strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString + "</b>"

                    strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block;'>" + rows.Item("DATA").ToString + "</span>"
                Else
                    Dim json As String = rows.Item("DATA").ToString
                    'PMHx,PSHx,Medications,FHx,Soc Hx,Ob Preg Hx,Hospitalizations,Allergies,Vitals
                    If rows.Item("NOTE SECTION TYPE").ToString = "PMHx" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString

                        For Each item As JProperty In data
                            item.CreateReader()
                            Dim sentance As String = ""
                            Select Case item.Name
                                Case "histories"
                                    For Each msg As JObject In item.Values
                                        Dim first_name As String = msg("blood_type")
                                        Dim last_name As String = msg("rh")
                                        Dim f As String = msg("hx_form").Children().ToList.Item(1).Children().ToList.Item(0)
                                        If String.IsNullOrEmpty(sentance) Then
                                            If String.IsNullOrEmpty(first_name) Then
                                                sentance = f + " " + first_name + "" + last_name + "<br/>"
                                            Else
                                                sentance = f + " " + first_name + "" + last_name + "<br/>"
                                            End If
                                        Else
                                            If String.IsNullOrEmpty(first_name) Then
                                                sentance = sentance + f + " " + first_name + "" + last_name + "<br/>"
                                            Else
                                                sentance = sentance + f + " " + first_name + "" + last_name + "<br/>"
                                            End If
                                        End If
                                    Next
                                Case "general_comments"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        'sentance = item.Children().ToList.Item(0)
                                        If String.IsNullOrEmpty(sentance) Then
                                            sentance = "General Comments: " + item.Children().ToList.Item(0).ToString + "<br/>"
                                        Else
                                            sentance = sentance + "General Comments: " + item.Children().ToList.Item(0).ToString + "<br/>"
                                        End If
                                    End If
                            End Select
                            strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                        Next
                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "PSHx" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "histories"
                                    For Each msg As JObject In item.Values
                                        Dim first_name As String = msg(" blood_type")
                                        Dim last_name As String = msg("rh")
                                        Dim f As String = msg("hx_form").Children().ToList.Item(1).Children().ToList.Item(0)
                                        Dim sentance As String = f + " " + first_name + "" + last_name
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                        'If Not String.IsNullOrEmpty(sentance) Then
                                        '    strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                        'Else
                                        '    strMessage = strMessage + "<br/>"
                                        'End If
                                    Next
                                Case "general_comments"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        Dim sentance As String = item.Children().ToList.Item(0).ToString
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; COLOR:black; FONT-WEIGHT: normal; display: block; margin-top: 4px; margin-bottom: 4px;'>General Comments: " + sentance + "</span>"
                                    End If
                            End Select
                        Next
                        'If strMessage.EndsWith("<br/>") Then
                        '    strMessage = strMessage + "<br/>"
                        'End If
                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "Medications" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                            Dim strFile As String = folderPath & DateTime.Today.ToString("dd-MMM-yyyy") & ".txt"
                            Dim fileExists As Boolean = File.Exists(strFile)
                            File.AppendAllText(strFile, dtVisit.Rows(0).Item("LAST NAME").ToString + ", " + dtVisit.Rows(0).Item("FIRST NAME").ToString + dtVisit.Rows(0).Item("NOTE_ID").ToString + $"Error in Med at-- {DateTime.Now}{Environment.NewLine}")
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString + "<br/>"
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "no_known_meds"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        Dim sentance As String = item.Children().ToList.Item(0).ToString
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                    End If
                                Case "medication_reco nciliation"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        Dim sentance As String = item.Children().ToList.Item(0).ToString
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; COLOR:black; FONT-WEIGHT: normal; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                    End If
                                Case "medication_reconciliation"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        Dim sentance As String = item.Children().ToList.Item(0).ToString
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; COLOR:black; FONT-WEIGHT: normal; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                    End If
                                Case "active_medications"
                                    Dim sentance As String = ""
                                    For Each items As JProperty In data
                                        items.CreateReader()
                                        Select Case items.Name
                                            Case "active_medications"
                                                For Each msg As JObject In items.Values
                                                    If Not String.IsNullOrEmpty(msg("raw_drug").ToString) Then
                                                        sentance = msg("raw_drug").ToString + ", Quanitiy: " + msg("quantity").ToString + " Sig: " + msg("sig").ToString + "<br/>"
                                                    Else
                                                        sentance = msg("raw_drug").ToString + ", Quanitiy: " + msg("quantity").ToString + " Sig: " + msg("sig").ToString + "<br/>"
                                                    End If


                                                Next
                                        End Select
                                    Next
                                    strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 2px; margin-bottom: 2px;'>" + sentance + "</span>"
                            End Select
                        Next

                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "FHx" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        If True Then
                            strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                            Dim sentance As String = ""
                            For Each item As JProperty In data
                                item.CreateReader()
                                Select Case item.Name
                                    Case "family_members"
                                        For Each msg As JObject In item.Values

                                            sentance = "Family Member: " + msg("member_type_name").ToString
                                            Dim last_name As String = ""
                                            If msg("is_alive") Then
                                                sentance = sentance + " is Alive "
                                            Else
                                                sentance = sentance + " is dead "
                                            End If

                                            If Not String.IsNullOrEmpty(msg("age_deceased")) Then
                                                sentance = sentance + ", Age Deceased " + msg("age_deceased").ToString
                                            End If
                                            Dim results1 As IList(Of JToken) = msg("family_histories").Children().ToList()
                                            Dim data1 As List(Of JToken) = results1.Children().ToList
                                            For Each items As JProperty In data1
                                                items.CreateReader()

                                                Select Case items.Name
                                                    Case "has_disease"
                                                        If items.Values.First Then
                                                            If sentance.Contains("has disease:") Then
                                                                sentance = sentance + " and "
                                                            Else
                                                                sentance = sentance + " has disease: "
                                                            End If

                                                        End If
                                                    Case "hx_form"
                                                        Dim results11 As IList(Of JToken) = items.Children().ToList()
                                                        Dim data11 As List(Of JToken) = results11.Children().ToList

                                                        For Each items1 As JProperty In data11
                                                            items1.CreateReader()

                                                            Select Case items1.Name
                                                                Case "name"
                                                                    sentance = sentance + " " + items1.Value.ToString

                                                            End Select
                                                        Next

                                                End Select
                                            Next


                                            strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                        Next
                                End Select
                            Next
                            'If Not String.IsNullOrEmpty(sentance) Then

                            'Else
                            '    strMessage = strMessage + "<br/>"
                            'End If

                        End If
                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "Soc Hx" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                        Dim sentance As String = ""
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "histories"
                                    Dim resultss As IList(Of JToken) = ser("histories").Children().ToList()
                                    Dim index As Int32 = resultss.Count
                                    For i = 0 To index - 1
                                        Dim results11 As IList(Of JToken) = item.Children().ToList.Item(0).Children.ToList.Item(i).Children.ToList
                                        Dim data11 As List(Of JToken) = results11.Children().ToList
                                        For Each itesm As JProperty In data11
                                            If True Then
                                                itesm.CreateReader()

                                                Select Case itesm.Name
                                                    Case "name"
                                                        sentance = sentance + " " + itesm.Value.ToString
                                                    Case "hx_form"
                                                        Dim results111 As IList(Of JToken) = itesm.Children().ToList()
                                                        Dim data112 As List(Of JToken) = results111.Children().ToList

                                                        For Each items1 As JProperty In data112
                                                            items1.CreateReader()

                                                            Select Case items1.Name
                                                                Case "name"
                                                                    'sentance = sentance + " " + items1.Value.ToString
                                                                    If String.IsNullOrEmpty(sentance) Then
                                                                        sentance = items1.Value.ToString
                                                                    Else
                                                                        sentance = sentance + ", " + items1.Value.ToString
                                                                    End If
                                                            End Select
                                                        Next
                                                    Case "comments"
                                                        If Not String.IsNullOrEmpty(itesm.Value.ToString) Then
                                                            sentance = sentance + "<br/>Comments: " + itesm.Value.ToString + "<br/>"
                                                        End If

                                                End Select
                                            End If
                                        Next
                                    Next

                                Case "general_comments"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then

                                        If String.IsNullOrEmpty(sentance) Then
                                            sentance = "<br/>General Comments: " + item.Children().ToList.Item(0).ToString + "<br/>"
                                        Else
                                            sentance = sentance + "<br/>General Comments: " + item.Children().ToList.Item(0).ToString + "<br/>"
                                        End If

                                    End If
                            End Select
                        Next
                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                        'If Not String.IsNullOrEmpty(sentance) Then

                        '    strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                        'Else
                        '    strMessage = strMessage + "<br/>"
                        'End If

                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "Ob Preg Hx" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try

                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                        Dim sentance As String = ""
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "ob_history"
                                    If item.Children().ToList.Item(0).Children.Count > 0 Then
                                        Dim resultss As IList(Of JToken) = item.Children().ToList.Item(0).Children.ToList.Item(3).Children.ToList
                                        Dim data111 As List(Of JToken) = resultss.Children().ToList
                                        Dim index As Int32 = data111.Count
                                        For i = 0 To index - 1
                                            Dim ser1 As JObject = JObject.Parse(data111.Item(i).ToString)
                                            Dim data11 As List(Of JToken) = ser1.Children().ToList
                                            Dim has_disease As String = ""
                                            Dim name As String = ""
                                            For Each itesm As JProperty In data11
                                                If True Then
                                                    itesm.CreateReader()

                                                    Select Case itesm.Name
                                                        Case "has_disease"
                                                            has_disease = itesm.Value.ToString
                                                        Case "hx_form"
                                                            Dim results111 As IList(Of JToken) = itesm.Children().ToList()
                                                            Dim data112 As List(Of JToken) = results111.Children().ToList

                                                            For Each items1 As JProperty In data112
                                                                items1.CreateReader()

                                                                Select Case items1.Name
                                                                    Case "name"
                                                                        name = items1.Value.ToString
                                                                End Select
                                                            Next
                                                        Case "comments"
                                                            If Not String.IsNullOrEmpty(itesm.Value.ToString) Then
                                                                sentance = sentance + "Comments:" + itesm.Value.ToString + "<br/>"
                                                            End If
                                                    End Select
                                                End If
                                            Next
                                            If has_disease = "False" Then
                                                If String.IsNullOrEmpty(sentance) Then
                                                    sentance = name + ": No<br/>"
                                                Else
                                                    sentance = sentance + "" + name + ": No<br/>"
                                                End If
                                            Else
                                                If String.IsNullOrEmpty(sentance) Then
                                                    sentance = name + ": Yes<br/>"
                                                Else
                                                    sentance = sentance + "" + name + ": Yes<br/>"
                                                End If
                                            End If
                                        Next
                                    End If
                                Case "ob_general_comments"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        If String.IsNullOrEmpty(sentance) Then
                                            sentance = "General Comments: " + item.Children().ToList.Item(0).ToString + "<br/>"
                                        Else
                                            sentance = sentance + "<br/>General Comments: " + item.Children().ToList.Item(0).ToString + "<br/>"
                                        End If
                                    End If


                            End Select
                        Next
                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "Hospitalizations" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "histories"
                                    For Each items As JProperty In data
                                        items.CreateReader()
                                        Select Case items.Name
                                            Case "histories"
                                                Dim sentance As String = ""
                                                For Each msg As JObject In items.Values

                                                    sentance = "Patient Admit on " + msg("admission_date").ToString + "<br/>" + "CPT Code(s): " + msg("cpt_code_text").ToString + "<br/>" + "Comments: " + msg("comments").ToString + "<br/>"

                                                Next
                                                strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 2px; margin-bottom: 2px;'>" + sentance + "</span>"
                                        End Select
                                    Next
                            End Select
                        Next
                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "Allergies" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try
                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "active_allergies"
                                    For Each items As JProperty In data
                                        items.CreateReader()
                                        Select Case items.Name
                                            Case "active_allergies"
                                                Dim sentance As String = ""
                                                For Each msg As JObject In items.Values
                                                    sentance = "Patient has " + msg("severity").ToString + " allergie with " + msg("name").ToString + " and has reaction " + msg("reaction").ToString + "<br/>"
                                                    If Not String.IsNullOrEmpty(msg("comments").ToString) Then
                                                        sentance = sentance + "Comments:" + msg("comments").ToString + "<br/>"
                                                    End If


                                                Next
                                                strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 2px; margin-bottom: 2px;'>" + sentance + "</span>"
                                        End Select
                                    Next
                                Case "no_known_allergies"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        Dim sentance As String = item.Children().ToList.Item(0).ToString
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; COLOR:black; FONT-WEIGHT: normal; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span><br/>"
                                    End If
                                Case "no_allergy_history"
                                    If Not String.IsNullOrEmpty(item.Children().ToList.Item(0).ToString) Then
                                        Dim sentance As String = item.Children().ToList.Item(0).ToString
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; COLOR:black; FONT-WEIGHT: normal; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span><br/>"
                                    End If
                            End Select
                        Next
                    ElseIf rows.Item("NOTE SECTION TYPE").ToString = "Vitals" Then
                        Dim ser As JObject = Nothing
                        Try
                            ser = JObject.Parse(json)
                        Catch ex As Exception
                            ser = JObject.Parse(json.Replace(" ", ""))
                        End Try

                        Dim data As List(Of JToken) = ser.Children().ToList
                        strMessage = strMessage + "<br/><br/><b style='FONT-SIZE:12pt; FONT-FAMILY:Arial; FONT-WEIGHT:bold; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + rows.Item("NOTE SECTION TYPE").ToString
                        Dim sentance As String = ""
                        For Each item As JProperty In data
                            item.CreateReader()
                            Select Case item.Name
                                Case "vitals"
                                    If item.Children().ToList.Item(0).Children.Count > 0 Then

                                        Dim ser11 As JObject = JObject.Parse(item.Children().ToList.Item(0).Children.ToList.Item(0).ToString)
                                        Dim data111 As List(Of JToken) = ser11.Children().ToList
                                        For Each items As JProperty In data111

                                            Dim has_disease As String = ""
                                            Dim name As String = ""
                                            items.CreateReader()

                                            Select Case items.Name
                                                Case "bmi"
                                                    If String.IsNullOrEmpty(sentance) Then
                                                        sentance = "BMI: " + items.Value.ToString + "<br/>"
                                                    Else
                                                        sentance = sentance + "BMI: " + items.Value.ToString + "<br/>"
                                                    End If
                                                Case "bp"
                                                    Dim results111 As IList(Of JToken) = items.Children().ToList()
                                                    Dim data112 As List(Of JToken) = results111.Children().ToList

                                                    For Each items1 As JProperty In data112
                                                        items1.CreateReader()

                                                        Select Case items1.Name
                                                            Case "measurement"
                                                                If String.IsNullOrEmpty(sentance) Then
                                                                    sentance = "Systolic BP is: " + items1.Value.ToString + "<br/>"
                                                                Else
                                                                    sentance = sentance + "Systolic BP is: " + items1.Value.ToString + "<br/>"
                                                                End If
                                                            Case "measurement2"
                                                                If String.IsNullOrEmpty(sentance) Then
                                                                    sentance = "Diastolic BP is: " + items1.Value.ToString + "<br/>"
                                                                Else
                                                                    sentance = sentance + "Diastolic BP is: " + items1.Value.ToString + "<br/>"
                                                                End If
                                                        End Select
                                                    Next
                                                Case "height"
                                                    If String.IsNullOrEmpty(sentance) Then
                                                        sentance = "Height is: " + items.Value.ToString + "<br/>"
                                                    Else
                                                        sentance = sentance + "Height is: " + items.Value.ToString + "<br/>"
                                                    End If
                                                Case "weight"
                                                    If String.IsNullOrEmpty(sentance) Then
                                                        sentance = "Weight is: " + items.Value.ToString + "<br/>"
                                                    Else
                                                        sentance = sentance + "Weight is: " + items.Value.ToString + "<br/>"
                                                    End If
                                                Case "comments"
                                                    If Not String.IsNullOrEmpty(items.Value.ToString) Then
                                                        sentance = sentance + "Comments:" + items.Value.ToString + "<br/>"
                                                    End If
                                            End Select
                                        Next
                                        strMessage = strMessage + "<br/><span style='FONT-SIZE:10pt; FONT-FAMILY:Arial; FONT-WEIGHT: normal; COLOR:black; display: block; margin-top: 4px; margin-bottom: 4px;'>" + sentance + "</span>"
                                    End If
                            End Select
                        Next
                    End If
                End If
            Next
            strMessage = strMessage + "</BODY></HTML>"
            Dim filepath As String = ""
            If dtVisit.Rows(0).Item("NOTE STATUS").ToString.StartsWith("com") Then
                filepath = folderPath + "\Compleate" & DateTime.Today.ToString("dd-MMM-yyyy")
            Else
                filepath = folderPath + "\InCompleate" & DateTime.Today.ToString("dd-MMM-yyyy")
            End If

            If Not Directory.Exists(filepath) Then
                Directory.CreateDirectory(filepath)
            End If
            Dim pdffileName As String = dtVisit.Rows(0).Item("FIRST NAME").ToString + "_" + dtVisit.Rows(0).Item("LAST NAME").ToString + "_" + dtVisit.Rows(0).Item("DATE OF BIRTH").ToString
            My.Computer.FileSystem.WriteAllText(filepath + "\" + pdffileName + ".html", strMessage, False)
            TmpTextControl = New ServerTextControl
            TmpTextControl.Create()
            TmpTextControl.Load(filepath + "\" + pdffileName + ".html", TXTextControl.StreamType.HTMLFormat)
            'If TmpTextControl.Text.Contains("Incompleate") Then
            '    TmpTextControl.Text = TmpTextControl.Text.Replace("(Incompleate)", "                    (Incompleate)")
            '    TmpTextControl.Text = TmpTextControl.Text.Replace("Frozen Lava Medical Center", "                    Frozen Lava Medical Center")
            'Else
            '    TmpTextControl.Text = TmpTextControl.Text.Replace("Frozen Lava Medical Center", "                    Frozen Lava Medical Center")
            'End If
            TmpTextControl.Save(filepath + "\" + pdffileName + ".pdf", StreamType.AdobePDF)
            TmpTextControl.Save(filepath + "\" + pdffileName + ".rtf", StreamType.RichTextFormat)
        Catch ex As Exception
            Dim strFile As String = folderPath & DateTime.Today.ToString("dd-MMM-yyyy") & ".txt"
            Dim fileExists As Boolean = File.Exists(strFile)
            File.AppendAllText(strFile, dtVisit.Rows(0).Item("LAST NAME").ToString + ", " + dtVisit.Rows(0).Item("FIRST NAME").ToString + dtVisit.Rows(0).Item("NOTE_ID").ToString + $"Error Message in  Occured at-- {DateTime.Now}{Environment.NewLine}")
        End Try

    End Sub

    Private Sub btnCreateTiff_Click(sender As Object, e As EventArgs) Handles btnCreateTiff.Click
        If String.IsNullOrEmpty(txtFileName.Text) Then
            Dim strRTFFilePath As String = "D:\\CCPCTiff.txt"
            Dim readText() As String = File.ReadAllLines(strRTFFilePath)
            For Each file As String In readText
                Using TmpTextcontrol As New TXTextControl.ServerTextControl
                    If Not TmpTextcontrol.Create Then
                        Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                    End If

                    TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(file & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)

                    Dim PageCount As Integer = 0
                    Dim inputImages As New ArrayList()
                    For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                        Dim image As New MemoryStream()
                        ' get the image from TX Text Control's page rendering engine
                        Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                        ' save and add the image to the ArrayList
                        mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                        inputImages.Add(image)
                        PageCount += 1
                    Next
                    appCtx.DocSvr.EDM.FileSystem.File.Move(file & "\ClinicalVisit.tif", file & "\ClinicalVisit-Bkp2.tif")

                    Decorator.DoWithTempFile(Sub(tmp)
                                                 CreateMultipageTIF(inputImages, tmp)
                                                 appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, file & "\ClinicalVisit.tif")
                                             End Sub)
                End Using
            Next
        Else
            Dim strRTFFilePath As String = txtFileName.Text
            ' File.Move(txtFileName.Text & "\ClinicalVisit.tif", txtFileName.Text & "\ClinicalVisit-Bkp.tif")
            Using TmpTextcontrol As New TXTextControl.ServerTextControl
                If Not TmpTextcontrol.Create Then
                    Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                End If

                'TmpTextcontrol.Load((strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
                TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)

                Dim PageCount As Integer = 0
                Dim inputImages As New ArrayList()
                For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                    Dim image As New MemoryStream()
                    ' get the image from TX Text Control's page rendering engine
                    Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                    ' save and add the image to the ArrayList
                    mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                    inputImages.Add(image)
                    PageCount += 1
                Next

                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.tif", strRTFFilePath & "\ClinicalVisit-Bkp2.tif")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.rtf", strRTFFilePath & "\ClinicalVisit-Bkp2.rtf")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.pdf", strRTFFilePath & "\ClinicalVisit-Bkp2.pdf")
                ''Creating new files tif, pdf, rtf
                Decorator.DoWithTempFile(Sub(tmp)
                                             CreateMultipageTIF(inputImages, tmp)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.tif")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.rtf")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.pdf")
                                         End Sub)


                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.tif", strRTFFilePath & "\ClinicalVisit-Bkp1.tif")

                Decorator.DoWithTempFile(Sub(tmp)
                                             CreateMultipageTIF(inputImages, tmp)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.tif")
                                         End Sub)
            End Using
            txtFileName.Text = ""
        End If

    End Sub

    Private Sub btnGetPlanIssue_Click(sender As Object, e As EventArgs) Handles btnGetPlanIssue.Click

        dtSectionUnit.Columns.Add("VISIT_SEQ_NUM")
        dtSectionUnit.Columns.Add("QUESTION_DESCRIPTION")
        dtSectionUnit.Columns.Add("SECTION_SEQ_NUM")
        dtSectionUnit.Columns.Add("VISIT_PATH")
        Dim Path As String = "PEHR6VisitList.txt"
        Dim readText() As String = File.ReadAllLines(Path)
        For i As Integer = 0 To readText.Length - 1
            Application.DoEvents()
            Dim visitpath As String = "C:\Users\mmasif\Downloads\200141328" 'readText(i)
            Label1.Text = "Processing " & (i + 1) & " of " & readText.Length
            If Not visitpath.Contains("Fixed") Then
                If GetPlanData(visitpath) Then
                    'My.Computer.FileSystem.WriteAllText(Path, My.Computer.FileSystem.ReadAllText(Path).Replace(visitpath, ""), False)
                    dtSectionUnit.WriteCSV("PEHR6PlanIssue.csv")
                End If
            End If
        Next
    End Sub

    Private Function GetPlanData(ByVal strvisitpath As String) As Boolean
        Dim visitID As String = Path.GetFileName(strvisitpath)
        Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(strvisitpath, "*.xml")
        If Not files.Empty() Then
            For Each fileName As String In files
                Try
                    Dim ds As New DataSet
                    If fileName.Contains("108948905166") OrElse fileName.Contains("10612335166") Then
                        Using ms As New MemoryStream(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(fileName))
                            ds.ReadXml(ms)
                        End Using
                        If ds.Tables.Contains("QUESTION_ROW") Then
                            For Each row As DataRow In ds.Tables("QUESTION_ROW").Select("DISPLAY_OBJECT = 'Procedure Control' and ANSWER_VALUE_UNIT is not null")
                                If row.Item("DISPLAY_OBJECT").ToString = "Procedure Control" AndAlso row.Item("ANSWER_VALUE").ToString = "N" Then
                                    Dim dom As New HtmlDocument
                                    dom.LoadHtml(ds.Tables("SECTION_ROW").Rows(0).Item("SECTION_HTML"))
                                    Dim Nodes As HtmlNodeCollection = dom.DocumentNode.SelectNodes("//span")
                                    If Nodes IsNot Nothing AndAlso Nodes.Count > 0 Then
                                        For Each Node As HtmlNode In Nodes
                                            Dim tmp As HtmlAttribute = Node.Attributes("id")
                                            If tmp.Value.ToString.Contains(row.Item("QUESTION_SEQ_NUM").ToString) Then
                                                If Not String.IsNullOrEmpty(Node.InnerText) AndAlso ds.Tables("SECTION_ROW").Row(0).Item("SECTION_HTML").ToString.Contains(Node.InnerText) Then
                                                    ' Dim strMatch As String = Node.InnerText
                                                    dtSectionUnit.Rows.Add(visitID, Node.InnerText, row.Item("SECTION_SEQ_NUM").ToString, strvisitpath)
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    End If
                Catch ex As Exception
                End Try
            Next
        End If

        Return True
    End Function
    Public maxID As Int32 = 0
    Public siglocation As Int32 = 0
    Dim ErrorList As String = "Start"
    Dim totalVisits As Integer = 0
    Private Sub btnHeaderSig_Click(sender As Object, e As EventArgs) Handles btnHeaderSig.Click

        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = sjfavisitString
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows

            If changeHeaderSig(rows) Then

            Else
                ErrorList = ErrorList + "---" + rows.Item("PATH").ToString + " Already fix or Encounter"
            End If
            Label1.Text = ""
            totalVisits = totalVisits + 1
            Label1.Text = totalVisits.ToString
        Next
        txtFileName.Text = ErrorList
    End Sub

    Private Function changeHeaderSig(ByVal strRTFFilePath As DataRow) As Boolean
        Try

            Dim visitID As String = ""
            Using TmpTextcontrol As New TXTextControl.ServerTextControl
                If Not TmpTextcontrol.Create Then
                    Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                End If
                Dim filepath As String = "D:\Currently working"
                If Not appCtx.DocSvr.EDM.FileSystem.File.Exists(strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.rtf") OrElse appCtx.DocSvr.EDM.FileSystem.File.Exists(strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp2.rtf") Then
                    Return False
                End If
                TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.rtf"), TXTextControl.StreamType.RichTextFormat)
                ''Header chnage logic
                Dim index As Integer = 0
                Dim header As TXTextControl.HeaderFooter = TmpTextcontrol.Sections(1).HeadersAndFooters.GetItem(TXTextControl.HeaderFooterType.Header)

                If header IsNot Nothing AndAlso header.TextFields IsNot Nothing Then
                    For Each fieldz As TextField In header.TextFields

                        If fieldz.Name.Contains("FIRST") Then
                            Return False
                        End If

                        If index = 0 Then
                            fieldz.Text = "SAN JUAN FOOT AND ANKLE CENTER"
                        ElseIf index = 1 Then
                            fieldz.Text = Environment.NewLine & strRTFFilePath.Item("PROVIDER_NAME")
                        ElseIf index = 2 Then
                            fieldz.Text = Environment.NewLine & "1825 EAST MAIN STREET, SUITE A"
                        ElseIf index = 3 Then
                            fieldz.Text = Environment.NewLine & "MONTROSE, CO 81401" & Environment.NewLine
                        End If
                        index = index + 1
                    Next
                End If

                ''Getting loation for Signature and field ID for custom feild 
                For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields
                    If tmpfield.ID = TmpTextcontrol.TextFields.Count - 4 Then
                        maxID = tmpfield.ID
                        siglocation = tmpfield.Length + tmpfield.Start
                    End If
                Next


                ''add signature field
                Dim bytes As Byte() = Nothing

                'Read signature
                bytes = File.ReadAllBytes("D:\Currently working\NewTemp\" + strRTFFilePath.Item("SIGNING_PROVIDER_SEQ_NUM").ToString & ".gif")
                If bytes IsNot Nothing Then
                    Dim strm As New MemoryStream(bytes)
                    Dim newImg As System.Drawing.Image = System.Drawing.Image.FromStream(strm)
                    Dim NewImage As New TXTextControl.Image(newImg)
                    NewImage.Sizeable = True
                    Dim bln As Boolean
                    bln = TmpTextcontrol.Images.Add(NewImage, siglocation)
                    If bln Then
                        NewImage.SaveMode = ImageSaveMode.SaveAsData
                    End If

                End If

                Dim dat As DateTime = strRTFFilePath.Item("SIGNATURE_TIMESTAMP")
                ''Adding custom field
                Dim Field As TXTextControl.TextField = CreateFieldCustomized(TmpTextcontrol.Selection, Name, "This visit was electronically signed off by " & strRTFFilePath.Item("PROVIDER_NAME") & " on " & dat.ToShortDateString & ".")

                Dim blnCheck As Boolean = TmpTextcontrol.TextFields.Add(Field)

                Dim PageCount As Integer = 0
                Dim inputImages As New ArrayList()
                For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                    Dim image As New MemoryStream()
                    Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                    mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                    inputImages.Add(image)
                    PageCount += 1
                Next
                ''Create backup of files
                'appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\" & visitID & "xmls.zip", strRTFFilePath.Item("PATH") & "\" & visitID & "xmls-Bkp1.zip")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp2.tif")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp2.rtf")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.pdf", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp2.pdf")
                ''Creating new files tif, pdf, rtf
                Decorator.DoWithTempFile(Sub(tmp)
                                             CreateMultipageTIF(inputImages, tmp)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.pdf")
                                         End Sub)
            End Using
            Return True
        Catch ex As Exception
            ErrorList = ErrorList + "---" + strRTFFilePath.Item("PATH").ToString + " has Issue"
            Return True
        End Try
    End Function
    Private Function CreateFieldCustomized(ByVal selection As TXTextControl.Selection, ByVal Name As String, ByVal Text As String) As TXTextControl.TextField
        Dim Field As New TXTextControl.TextField
        selection.Start = 0
        selection.Length = Integer.MaxValue

        selection.ForeColor = System.Drawing.Color.Black
        'selection.Bold = False
        'selection.Underline = FontUnderlineStyle.None
        selection.FontSize = 200
        selection.FontName = "Arial"


        Field.Name = Name
        Field.Text = Environment.NewLine & Text
        Field.DoubledInputPosition = True
        Field.Deleteable = False
        Field.Editable = False
        Field.ShowActivated = False

        Field.ID = maxID + 1

        Return Field
    End Function
    Private Function CreateCPTField(ByVal selection As TXTextControl.Selection, ByVal Name As String, ByVal Text As String, ByVal ID As Integer) As TXTextControl.TextField
        Dim Field As New TXTextControl.TextField
        selection.Start = 0
        selection.Length = Integer.MaxValue

        selection.ForeColor = System.Drawing.Color.Black
        'selection.Bold = False
        'selection.Underline = FontUnderlineStyle.None
        selection.FontSize = 200
        selection.FontName = "Arial"


        Field.Name = Name
        Field.Text = Environment.NewLine & Text
        Field.DoubledInputPosition = True
        Field.Deleteable = False
        Field.Editable = False
        Field.ShowActivated = False

        Field.ID = ID

        Return Field
    End Function

    Private Sub InsertImage(ByVal SigProvSeqNum As String, ByVal InputPosition As Integer, ByRef objHeaderFooter As TXTextControl.HeaderFooter)
        If SigProvSeqNum <> "" Then
            Dim bytes As Byte() = Nothing

            'Read signature
            'FileService.Instance.ReadSignatureFile(SigProvSeqNum, bytes, "\\" + _EDmHost + "\" + _EDmAlias)

            If bytes IsNot Nothing Then
                Dim strm As New MemoryStream(bytes)
                Dim newImg As System.Drawing.Image = System.Drawing.Image.FromStream(strm)
                Dim NewImage As New TXTextControl.Image(newImg)
                NewImage.Sizeable = True
                Dim bln As Boolean
                If objHeaderFooter IsNot Nothing Then
                    bln = objHeaderFooter.Images.Add(NewImage, 5)
                Else
                    '  bln = TmpTextcontrol.Images.Add(NewImage, New Point(5000, 5567), ImageInsertionMode.DisplaceText)
                    bln = TmpTextControl.Images.Add(NewImage, 5)
                End If
                If bln Then
                    NewImage.SaveMode = ImageSaveMode.SaveAsData
                End If

            End If
        End If

    End Sub

    Private Sub btnPastVisitFileCopy_Click(sender As Object, e As EventArgs) Handles btnPastVisitFileCopy.Click
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = ccpcvisitString
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows
            'rows.Item("TO_COPY")
            'appCtx.DocSvr.EDM.FileSystem.File.Copy("from", "to", False)
            ' GetSectionData(rows.Item("TO_COPY").ToString)
            Dim visitID As String = Path.GetFileName(rows.Item("TO_COPY").ToString)
            Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(rows.Item("FROM_COPY").ToString, "*.xml", [option]:=SearchOption.AllDirectories)

            If Not files.Empty() Then
                For Each fileName As String In files
                    Try
                        If Not fileName.Contains("PickList") AndAlso Not fileName.Contains("27.xml") AndAlso Not fileName.Contains("28.xml") Then
                            'Try
                            '    If appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(rows.Item("TO_COPY") + "\" + Path.GetFileName(fileName)).Count < 50 Then
                            '        appCtx.DocSvr.EDM.FileSystem.File.Copy(rows.Item("FROM_COPY") + "\" + Path.GetFileName(fileName), rows.Item("TO_COPY") + "\" + Path.GetFileName(fileName), True)
                            '    End If
                            'Catch ex As Exception
                            '    appCtx.DocSvr.EDM.FileSystem.File.Copy(rows.Item("FROM_COPY") + "\" + Path.GetFileName(fileName), rows.Item("TO_COPY") + "\" + Path.GetFileName(fileName), True)
                            'End Try
                            appCtx.DocSvr.EDM.FileSystem.File.Copy(rows.Item("FROM_COPY") + "\" + Path.GetFileName(fileName), rows.Item("TO_COPY") + "\" + Path.GetFileName(fileName), True)
                        End If
                    Catch ex As Exception
                    End Try
                Next
            End If
        Next
        txtFileName.Text = ErrorList
    End Sub

    Private Sub btnSWHeaderimg_Click(sender As Object, e As EventArgs) Handles btnSWHeaderimg.Click

        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = SWvisitString
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows

            If SWchangeHeaderSig(rows) Then

            Else
                ErrorList = ErrorList + "---" + rows.Item("PATH").ToString + " Already fix or Encounter"
                My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", rows.Item("PATH").ToString + " has Issue", True)
            End If
            Label1.Text = ""
            totalVisits = totalVisits + 1
            Label1.Text = totalVisits.ToString
        Next
        txtFileName.Text = ErrorList
    End Sub

    Private Function SWchangeHeaderSig(ByVal strRTFFilePath As DataRow) As Boolean
        Try

            Dim visitID As String = ""
            Using TmpTextcontrol As New TXTextControl.ServerTextControl
                If Not TmpTextcontrol.Create Then
                    Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                End If
                Dim filepath As String = "D:\Currently working"
                'If appCtx.DocSvr.EDM.FileSystem.File.Exists(strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.rtf") OrElse appCtx.DocSvr.EDM.FileSystem.File.Exists(strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp2.rtf") Then
                '    Return False
                'End If
                TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
                ''Header chnage logic
                Dim index As Integer = 0
                Dim header As TXTextControl.HeaderFooter = TmpTextcontrol.Sections(1).HeadersAndFooters.GetItem(TXTextControl.HeaderFooterType.Header)

                If header IsNot Nothing AndAlso header.TextFields IsNot Nothing Then
                    For Each fieldz As TextField In header.TextFields
                        If Not String.IsNullOrEmpty(fieldz.Text) Then
                            fieldz.Text = ""
                        End If
                    Next
                    Dim header1 As TXTextControl.HeaderFooter = TmpTextcontrol.Sections(1).HeadersAndFooters.GetItem(TXTextControl.HeaderFooterType.Header)
                    If header1 IsNot Nothing AndAlso header1.TextFields IsNot Nothing Then
                        For Each fieldz1 As TextField In header1.TextFields
                            If Not String.IsNullOrEmpty(fieldz1.Text) Then
                                fieldz1.Text = ""
                            End If
                        Next
                    End If
                    Dim header2 As TXTextControl.HeaderFooter = TmpTextcontrol.Sections(1).HeadersAndFooters.GetItem(TXTextControl.HeaderFooterType.Header)

                    If header2 IsNot Nothing AndAlso header2.TextFields IsNot Nothing Then
                        For Each fieldz2 As TextField In header2.TextFields
                            If Not String.IsNullOrEmpty(fieldz2.Text) Then
                                fieldz2.Text = ""
                            End If


                        Next
                    End If
                End If

                Dim bytes As Byte() = Nothing

                'Read Header
                bytes = File.ReadAllBytes("D:\Currently working\NewTemp\Header.png")
                If bytes IsNot Nothing Then
                    Dim strm As New MemoryStream(bytes)
                    Dim newImg As System.Drawing.Image = System.Drawing.Image.FromStream(strm)
                    Dim NewImage As New TXTextControl.Image(newImg)
                    NewImage.Sizeable = True
                    Dim bln As Boolean
                    bln = TmpTextcontrol.Images.Add(NewImage, 0)
                    If bln Then
                        NewImage.SaveMode = ImageSaveMode.SaveAsData
                    End If

                End If

                ''Getting loation for Signature and field ID for custom feild 
                For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields
                    If tmpfield.ID = TmpTextcontrol.TextFields.Count - 4 Then
                        maxID = tmpfield.ID
                        siglocation = tmpfield.Length + tmpfield.Start
                    End If
                Next


                Dim dat As DateTime = strRTFFilePath.Item("SIGNATURE_TIMESTAMP")
                ''Adding custom field
                Dim Field As TXTextControl.TextField = CreateFieldCustomized(TmpTextcontrol.Selection, Name, "This visit was electronically signed off by " & strRTFFilePath.Item("PROVIDER_NAME") & " on " & dat.ToShortDateString & ".")

                Dim blnCheck As Boolean = TmpTextcontrol.TextFields.Add(Field)

                Dim PageCount As Integer = 0
                Dim inputImages As New ArrayList()
                For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                    Dim image As New MemoryStream()
                    Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                    mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                    inputImages.Add(image)
                    PageCount += 1
                Next

                ''Create backup of files
                'appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\" & visitID & "xmls.zip", strRTFFilePath.Item("PATH") & "\" & visitID & "xmls-Bkp1.zip")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.tif")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.rtf")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.pdf", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.pdf")
                ''Creating new files tif, pdf, rtf

                Decorator.DoWithTempFile(Sub(tmp)
                                             CreateMultipageTIF(inputImages, tmp)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.pdf")
                                         End Sub)
            End Using
            Return True
        Catch ex As Exception
            ErrorList = ErrorList + "---" + strRTFFilePath.Item("PATH").ToString + " has Issue"
            My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", strRTFFilePath.Item("PATH").ToString + " has Issue", True)
            Return True
        End Try
    End Function

    Private Sub btnSDFAPastVisitFix_Click(sender As Object, e As EventArgs) Handles btnSDFAPastVisitFix.Click
        Dim strRTFFilePath As String = "D:\\SDFAVisits.txt"
        Dim readText() As String = File.ReadAllLines(strRTFFilePath)
        For Each file As String In readText
            If SDFAvisitfix(file) Then
                My.Computer.FileSystem.WriteAllText(strRTFFilePath, My.Computer.FileSystem.ReadAllText(strRTFFilePath).Replace(file, ""), False)
            Else
                ErrorList = ErrorList + "---" + file + " Already fix or Encounter"
                My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", file + " has Issue", True)
            End If

        Next

    End Sub
    Private Function SDFAvisitfix(ByVal strRTFFilePath As String) As Boolean
        Try
            Dim visitpath As String = strRTFFilePath
            Dim visitID As String = Path.GetFileName(visitpath)
            Dim lastProtocol As SecurityProtocolType = System.Net.ServicePointManager.SecurityProtocol
            System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType)
            Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(visitpath)
            If Not files.Empty() Then
                For Each fileName As String In files
                    Try
                        If fileName.Contains("1000003905708") Then
                            'appCtx.DocSvr.EDM.FileSystem.File.Move(visitpath & "\1000003905708.xml", visitpath & "\1000003905708-backup.xml")
                            appCtx.DocSvr.EDM.FileSystem.File.Copy(visitpath & "\1000003905708.xml", visitpath & "\1000003905708-backup.xml")
                            Dim xmlDoc As Xml.XmlDocument = Nothing, xmlRoot As Xml.XmlNode = Nothing
                            Dim xmlNode As Xml.XmlNode = Nothing, xmlChild As Xml.XmlNode = Nothing

                            Try
                                xmlDoc = New Xml.XmlDocument()
                                ' xmlDoc.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllText(fileName))
                                xmlDoc.LoadXml(appCtx.DocSvr.EDM.FileSystem.File.ReadAllText(fileName))
                                xmlRoot = xmlDoc.DocumentElement()
                                xmlNode = xmlRoot.ChildNodes(0)
                                For Index = 1 To 4
                                    For Each xmlChild In xmlNode
                                        'MsgBox(xmlChild.Name & vbCr & xmlChild.InnerText)
                                        If Not xmlChild.Name = "SECTION_HTML" Then
                                            If xmlChild.InnerText.Contains("SEQUEL PAST MEDICATIONS") Then
                                                xmlChild.ParentNode.RemoveChild(xmlChild)
                                            End If
                                            If xmlChild.InnerText.Contains("SEQUEL ALLERGIES") Then
                                                xmlChild.ParentNode.RemoveChild(xmlChild)
                                            End If
                                            If xmlChild.InnerText.Contains("SEQUEL PROBLEM LIST") Then
                                                xmlChild.ParentNode.RemoveChild(xmlChild)
                                            End If
                                            If xmlChild.InnerText.Contains("SEQUEL CURRENT MEDICATIONS") Then
                                                xmlChild.ParentNode.RemoveChild(xmlChild)
                                            End If
                                        End If
                                    Next
                                Next

                                Dim xmlStream As MemoryStream = New MemoryStream()
                                xmlDoc.Save(xmlStream)
                                appCtx.DocSvr.EDM.FileSystem.File.WriteAllBytes(fileName, xmlStream.ToArray)
                                xmlStream.Flush()
                                xmlStream.Position = 0

                                Dim intZipCountr As Integer
                                Dim slFileList As New SortedList

                                For Each objFile As Object In files
                                    If objFile.ToString.Contains(".xml") OrElse objFile.ToString.Contains(".rtf") Then
                                        If Not objFile.ToString.Contains("backup") Then
                                            slFileList.Add(intZipCountr, objFile)
                                            intZipCountr += 1
                                        End If
                                    End If

                                Next
                                appCtx.DocSvr.EDM.FileSystem.File.Move(visitpath & "\" & visitID & "xmls.zip", visitpath & "\" & visitID & "xmls-Bkp1.zip")
                                'appCtx.DocSvr.EDM.HttpFs.Zip.CreateZip(slFileList, visitpath & "\" & visitID & "xmls.zip")
                                'appCtx.DocSvr.EDM.FileSystem.Zip.CreateZip(slFileList, visitpath & "\" & visitID & "xmls.zip")
                                Return True
                            Catch ex As Exception
                                ErrorList = ErrorList + "---" + strRTFFilePath + " has Issue"
                                My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", strRTFFilePath + " has Issue", True)
                                Return True
                            Finally
                                xmlRoot = Nothing
                                xmlNode = Nothing
                                xmlDoc = Nothing
                            End Try
                        End If
                        System.Net.ServicePointManager.SecurityProtocol = lastProtocol
                    Catch ex As Exception
                        ErrorList = ErrorList + "---" + strRTFFilePath + " has Issue"
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", strRTFFilePath + " has Issue", True)
                        Return True
                    End Try
                Next
            End If
            Return True
        Catch ex As Exception
            ErrorList = ErrorList + "---" + strRTFFilePath + " has Issue"
            My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", strRTFFilePath + " has Issue", True)
            Return True
        End Try
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = ccpcvisitString.Replace("GRANDCC997", "112233")
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows

            If SectionSig(rows) Then

            Else
                ErrorList = ErrorList + "---" + rows.Item("PATH").ToString + " Already fix or Encounter"
                My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & rows.Item("PATH").ToString + " has Issue", True)
            End If
            'Label1.Text = ""
            totalVisits = totalVisits + 1
            Label1.Text = totalVisits.ToString
        Next
        txtFileName.Text = ErrorList
    End Sub
    Dim IsChange As Boolean = False
    Private Function RemoveTagsFromHTMl(ByVal str As String) As String
        Dim wc As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(str, "Seperator="".*?""", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        For Each match As System.Text.RegularExpressions.Match In wc
            str = str.Replace(match.Value, "")
        Next
        wc = System.Text.RegularExpressions.Regex.Matches(str, "Ending="".*?""", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        For Each match As System.Text.RegularExpressions.Match In wc
            str = str.Replace(match.Value, "")
        Next
        wc = System.Text.RegularExpressions.Regex.Matches(str, "Conjunction="".*?""", System.Text.RegularExpressions.RegexOptions.Singleline Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        For Each match As System.Text.RegularExpressions.Match In wc
            str = str.Replace(match.Value, "Conjunction="".""")
        Next
        Return str
    End Function
    Private Function SectionSig(ByVal strRTFFilePath As DataRow) As Boolean
        IsChange = False
        Dim visitpath As String = strRTFFilePath.Item("PATH")
        Dim visitID As String = Path.GetFileName(visitpath)
        Try
            Dim lastProtocol As SecurityProtocolType = System.Net.ServicePointManager.SecurityProtocol
            System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType)
            'Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(visitpath)
            'If Not files.Empty() Then
            Try
                Using TmpTextcontrol As New TXTextControl.ServerTextControl
                    If Not TmpTextcontrol.Create Then
                        Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                    End If
                    TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(visitpath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
                    My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\DetailLog.txt", Environment.NewLine & visitID + " Process Started", True)
                    '
                    'For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields

                    '    ''Adding Missing CPTs
                    '    Dim strdata As String = tmpfield.Text
                    '    Dim strolddata As String = ""
                    '    Dim StrCPTs As String = ""
                    '    For Each fileName As String In files
                    '        Dim ds As New DataSet
                    '        If Not fileName.Contains("PickList") AndAlso (fileName.Contains("13795207") OrElse fileName.Contains("13955207")) Then
                    '            If strdata.StartsWith("Plan") Then
                    '                Using ms As New MemoryStream(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(fileName))
                    '                    ds.ReadXml(ms)
                    '                End Using
                    '                siglocation = tmpfield.Length + tmpfield.Start + 1
                    '                Dim DrQuestions As DataRow() = Nothing
                    '                If fileName.Contains("13795207") Then
                    '                    DrQuestions = ds.Tables("QUESTION_ROW").Select("PARENT_HEADING_ID = " & ds.Tables("HEADING_ROW").Select("DESCRIPTION Like 'Treatment Provided%'")(0).Item("HEADING_ID").ToString)
                    '                Else fileName.Contains("13955207")
                    '                    DrQuestions = ds.Tables("QUESTION_ROW").Select("PARENT_HEADING_ID = " & ds.Tables("HEADING_ROW").Select("DESCRIPTION Like 'Procedures%'")(0).Item("HEADING_ID").ToString)
                    '                End If

                    '                If DrQuestions IsNot Nothing AndAlso DrQuestions.Count > 0 Then
                    '                    For Each dr As DataRow In DrQuestions
                    '                        If Not String.IsNullOrEmpty(dr.Item("ANSWER_VALUE").ToString) AndAlso dr.Item("ANSWER_VALUE") = "Y" Then
                    '                            Dim strval As String = dr.Item("DESCRIPTION")
                    '                            If Not strdata.Contains(strval) Then
                    '                                If String.IsNullOrEmpty(StrCPTs) Then
                    '                                    StrCPTs = strval
                    '                                Else
                    '                                    StrCPTs = StrCPTs & "<br>" & strval
                    '                                End If
                    '                            Else
                    '                                strolddata = strval
                    '                            End If
                    '                        End If
                    '                    Next
                    '                    If Not String.IsNullOrEmpty(StrCPTs) Then

                    '                        Dim strsectionhtml As String = ds.Tables("SECTION_ROW").Rows(0).Item("SECTION_HTML").ToString
                    '                        If Not String.IsNullOrEmpty(strolddata) Then
                    '                            strsectionhtml = strsectionhtml.Replace(strolddata, strolddata & "<br>" & StrCPTs)
                    '                        Else
                    '                            Dim font1 As System.Drawing.Font = New System.Drawing.Font("Arial", 10.0!)

                    '                            Dim pos As Integer = strsectionhtml.LastIndexOf("<")
                    '                            If pos <> -1 Then
                    '                                strsectionhtml = strsectionhtml.Substring(0, pos)
                    '                                strsectionhtml = strsectionhtml & "<span" + " style='FONT-WEIGHT: normal; FONT-SIZE: " + font1.SizeInPoints.ToString + "pt; FONT-FAMILY: " + font1.Name + "'" + ">" + StrCPTs + "</span></span>"
                    '                            End If
                    '                        End If

                    '                        strsectionhtml = strsectionhtml.Replace("blue", "black")
                    '                        Dim sectionSeqNum As String = ds.Tables("SECTION_ROW").Rows(0).Item("SECTION_SEQ_NUM").ToString
                    '                        Dim eleSection As mshtml.HTMLSpanElement

                    '                        Dim _HTMLDocument As mshtml.HTMLDocument = Nothing
                    '                        Dim htmlDoc As mshtml.IHTMLDocument2 = CType(New mshtml.HTMLDocument, mshtml.IHTMLDocument2)
                    '                        htmlDoc.write(strsectionhtml)
                    '                        _HTMLDocument = New mshtml.HTMLDocument()
                    '                        _HTMLDocument = CType(htmlDoc, mshtml.HTMLDocument)
                    '                        eleSection = CType(_HTMLDocument.getElementById(sectionSeqNum), mshtml.HTMLSpanElement)

                    '                        TmpTextcontrol.Selection.Start = tmpfield.Start - 1
                    '                        TmpTextcontrol.Selection.Length = tmpfield.Length
                    '                        Dim sectionHtml As String = RemoveTagsFromHTMl(eleSection.outerHTML.ToString)
                    '                        Dim font As System.Drawing.Font = New System.Drawing.Font("Arial", 10.0!)
                    '                        sectionHtml = "<SPAN" + " style='FONT-WEIGHT: normal; FONT-SIZE: " + font.SizeInPoints.ToString + "pt; FONT-FAMILY: " + font.Name + "'" + ">" + sectionHtml + "</SPAN>"
                    '                        'TODO: Need to review, fonts are configureable in SequelMedEMR
                    '                        TmpTextcontrol.Selection.Load(sectionHtml, TXTextControl.StringStreamType.HTMLFormat)
                    '                        TmpTextcontrol.Selection.Start = tmpfield.Start - 1
                    '                        TmpTextcontrol.Selection.Length = tmpfield.Length
                    '                        Try
                    '                            TmpTextcontrol.Selection.Save(sectionHtml, StringStreamType.HTMLFormat)
                    '                        Catch ex As Exception

                    '                        End Try

                    '                        IsChange = True
                    '                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\FixingLog.txt", Environment.NewLine & visitID + " Data added " + StrCPTs, True)
                    '                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\CPTAdded.txt", Environment.NewLine & visitID + "- " + StrCPTs, True)
                    '                    End If
                    '                End If
                    '            End If
                    '        End If
                    '    Next
                    '    If tmpfield.ID = TmpTextcontrol.TextFields.Count - 4 Then
                    '        maxID = tmpfield.ID
                    '    End If
                    '    If tmpfield.Name.Contains("PROVIDER_FIRST_NAME") AndAlso tmpfield.Text.ToUpper <> strRTFFilePath.Item("FIRST_NAME") Then
                    '        tmpfield.Text = strRTFFilePath.Item("FIRST_NAME")
                    '        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\FixingLog.txt", Environment.NewLine & visitID + " Provider First Name Change", True)
                    '        IsChange = True
                    '    End If

                    '    If tmpfield.Name.Contains("PROVIDER_LAST_NAME") AndAlso tmpfield.Text.ToUpper <> strRTFFilePath.Item("LAST_NAME") Then
                    '        tmpfield.Text = strRTFFilePath.Item("LAST_NAME")
                    '        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\FixingLog.txt", Environment.NewLine & visitID + " Provider Last Name Change", True)
                    '        IsChange = True
                    '    End If

                    '    If tmpfield.Name.Contains("PROVIDER_QUALIFICATION") AndAlso tmpfield.Text.ToUpper <> strRTFFilePath.Item("QUALIFICATION") Then
                    '        tmpfield.Text = strRTFFilePath.Item("QUALIFICATION")
                    '        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\FixingLog.txt", Environment.NewLine & visitID + " Provider Qualification Change", True)
                    '        IsChange = True
                    '    End If
                    'Next

                    'If TmpTextcontrol.Text.ToUpper.Contains("THIS VISIT WAS ELECTRONICALLY SIGNED OFF BY HENRY HALL") Then
                    '    My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\provider.txt", Environment.NewLine & visitID + "  HENRY HALL Visits", True)
                    'End If

                    ''Addind Signoff line
                    If Not TmpTextcontrol.Text.Contains("This visit was electronically signed off by") Then

                        Dim dat As DateTime = strRTFFilePath.Item("SIGNATURE_TIMESTAMP")
                        ''Adding custom field
                        Dim Field As TXTextControl.TextField = CreateFieldCustomized(TmpTextcontrol.Selection, Name, "This visit was electronically signed off by " & strRTFFilePath.Item("PROVIDER_NAME") & " on " & dat.ToShortDateString & ".")

                        Dim blnCheck As Boolean = TmpTextcontrol.TextFields.Add(Field)
                        IsChange = True
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\FixingLog.txt", Environment.NewLine & visitID + " Sign off Line Added", True)
                    End If

                    If IsChange Then
                        Dim PageCount As Integer = 0
                        Dim inputImages As New ArrayList()
                        For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                            Dim image As New MemoryStream()
                            Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                            mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                            inputImages.Add(image)
                            PageCount += 1
                        Next

                        ' Dim localPath As String = "D:\visit fixing"
                        'TmpTextcontrol.Save(localPath & "\ClinicalVisit.pdf", TXTextControl.StreamType.AdobePDF)
                        ''Create backup of files
                        appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.tif")
                        appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.rtf")
                        appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\ClinicalVisit.pdf", strRTFFilePath.Item("PATH") & "\ClinicalVisit-Bkp1.pdf")
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\DetailLog.txt", Environment.NewLine & visitID + " Backup Done", True)
                        ''Creating new files tif, pdf, rtf
                        Decorator.DoWithTempFile(Sub(tmp)
                                                     CreateMultipageTIF(inputImages, tmp)
                                                     appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif")
                                                     TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                                     appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf")
                                                     TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                                     appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.pdf")
                                                 End Sub)
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\DetailLog.txt", Environment.NewLine & visitID + " Process Ended", True)
                        IsChange = False
                    Else
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & visitID + " No Such issue found", True)
                    End If
                End Using
            Catch ex As Exception
                ErrorList = ErrorList + "---" + visitID + " has Issue"
                My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & visitID + " has Issue", True)
                Return True
            End Try

            ''End If
            System.Net.ServicePointManager.SecurityProtocol = lastProtocol
            IsChange = False
            Return True
        Catch ex As Exception
            ErrorList = ErrorList + "---" + visitID + " has Issue"
            My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & visitID + " has Issue", True)
            Return True
        End Try
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not String.IsNullOrEmpty(txtFileName.Text) Then
            changeHeaderSig1(txtFileName.Text)
            txtFileName.Text = ""
        End If
    End Sub
    Private Function changeHeaderSig1(ByVal strRTFFilePath As String) As Boolean
        Try

            Dim visitID As String = ""
            Using TmpTextcontrol As New TXTextControl.ServerTextControl
                If Not TmpTextcontrol.Create Then
                    Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                End If
                Dim filepath As String = "D:\Currently working"
                'If Not appCtx.DocSvr.EDM.FileSystem.File.Exists(strRTFFilePath & "\ClinicalVisit-Bkp1.rtf") OrElse appCtx.DocSvr.EDM.FileSystem.File.Exists(strRTFFilePath & "\ClinicalVisit-Bkp2.rtf") Then
                '    Return False
                'End If
                TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
                ''Header chnage logic

                Dim header As TXTextControl.HeaderFooter = TmpTextcontrol.Sections(1).HeadersAndFooters.GetItem(TXTextControl.HeaderFooterType.Header)

                If header IsNot Nothing AndAlso header.TextFields IsNot Nothing Then
                    For Each fieldz As TextField In header.TextFields
                        'TIPI
                        'If fieldz.Name.Contains("LOCATION_DESCRIPTION") Then
                        '    fieldz.Text = "TENNESSEE INTEGRATIVE PAIN INSTITUTE"
                        'End If
                        'If fieldz.Name.Contains("LOCATION_ADDRESS1") Then
                        '    fieldz.Text = "5073 MAIN STREET, STE 200"
                        'End If
                        'If fieldz.Name.Contains("LOCATION_CITY") Then
                        '    fieldz.Text = "SPRING HILL"
                        'End If
                        'If fieldz.Name.Contains("LOCATION_STATE") Then
                        '    fieldz.Text = "TN"
                        'End If
                        'If fieldz.Name.Contains("LOCATION_ZIP") Then
                        '    fieldz.Text = "37174"
                        'End If
                        'If fieldz.Name.Contains("LOCATION_TEL_NUM1") Then
                        '    fieldz.Text = "(615)206-3014"
                        'End If
                        'TIH
                        If fieldz.Name.Contains("LOCATION_DESCRIPTION") Then
                            fieldz.Text = "TENNESSEE INTEGRATIVE HEALTHCARE"
                        End If
                        If fieldz.Name.Contains("LOCATION_ADDRESS1") Then
                            fieldz.Text = "5055 MARYLAND WAY, STE 100"
                        End If
                        If fieldz.Name.Contains("LOCATION_CITY") Then
                            fieldz.Text = "BRENTWOOD"
                        End If
                        If fieldz.Name.Contains("LOCATION_STATE") Then
                            fieldz.Text = "TN"
                        End If
                        If fieldz.Name.Contains("LOCATION_ZIP") Then
                            fieldz.Text = "37027"
                        End If
                        If fieldz.Name.Contains("LOCATION_TEL_NUM1") Then
                            fieldz.Text = "(253)414-3066"
                        End If
                    Next
                End If

                For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields

                    If tmpfield.Name.Contains("PROVIDER_FIRST_NAME") Then

                    End If
                Next

                Dim PageCount As Integer = 0
                Dim inputImages As New ArrayList()
                For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                    Dim image As New MemoryStream()
                    Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                    mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                    inputImages.Add(image)
                    PageCount += 1
                Next
                ''Create backup of files
                'appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("PATH") & "\" & visitID & "xmls.zip", strRTFFilePath.Item("PATH") & "\" & visitID & "xmls-Bkp1.zip")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.tif", strRTFFilePath & "\ClinicalVisit-Bkp1.tif")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.rtf", strRTFFilePath & "\ClinicalVisit-Bkp1.rtf")
                appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.pdf", strRTFFilePath & "\ClinicalVisit-Bkp1.pdf")
                ''Creating new files tif, pdf, rtf
                Decorator.DoWithTempFile(Sub(tmp)
                                             CreateMultipageTIF(inputImages, tmp)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.tif")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.rtf")
                                             TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                             appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.pdf")
                                         End Sub)
            End Using
            txtFileName.Text = ""
            Return True
        Catch ex As Exception
            ErrorList = ErrorList + "---" + strRTFFilePath + " has Issue"
            Return True
        End Try
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim visitpath As String = txtFileName.Text 'strRTFFilePath.Item("PATH")
        Dim visitID As String = Path.GetFileName(visitpath)
        Try
            Dim lastProtocol As SecurityProtocolType = System.Net.ServicePointManager.SecurityProtocol
            System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType)

            Try
                Using TmpTextcontrol As New TXTextControl.ServerTextControl
                    If Not TmpTextcontrol.Create Then
                        Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                    End If
                    TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(visitpath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
                    My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\DetailLog.txt", Environment.NewLine & visitID + " Process Started", True)
                    '
                    For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields

                        If tmpfield.Name.Contains("PROVIDER_FIRST_NAME") Then
                            tmpfield.Text = "ASHLEY"
                            IsChange = True
                        End If

                        If tmpfield.Name.Contains("PROVIDER_LAST_NAME") Then
                            tmpfield.Text = "GLASER"
                            IsChange = True
                        End If

                        If tmpfield.Name.Contains("PROVIDER_QUALIFICATION") Then
                            tmpfield.Text = "FNP-BC"
                            IsChange = True
                        End If
                    Next
                    If TmpTextcontrol.Text.Contains("HENRY") Then
                        TmpTextcontrol.Text.Replace("HENRY", "NELSON")
                        IsChange = True
                    End If
                    If TmpTextcontrol.Text.Contains("HALL") Then
                        TmpTextcontrol.Text.Replace("HALL", "HALL")
                        IsChange = True
                    End If
                    If TmpTextcontrol.Text.Contains("DC") Then
                        TmpTextcontrol.Text.Replace("DC", "DC # X012986-1")
                        IsChange = True
                    End If
                    If IsChange Then
                        Dim PageCount As Integer = 0
                        Dim inputImages As New ArrayList()
                        For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                            Dim image As New MemoryStream()
                            Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                            mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                            inputImages.Add(image)
                            PageCount += 1
                        Next

                        ' Dim localPath As String = "D:\visit fixing"
                        'TmpTextcontrol.Save(localPath & "\ClinicalVisit.pdf", TXTextControl.StreamType.AdobePDF)
                        ''Create backup of files
                        appCtx.DocSvr.EDM.FileSystem.File.Move(visitpath & "\ClinicalVisit.tif", visitpath & "\ClinicalVisit-Bkp1.tif")
                        appCtx.DocSvr.EDM.FileSystem.File.Move(visitpath & "\ClinicalVisit.rtf", visitpath & "\ClinicalVisit-Bkp1.rtf")
                        appCtx.DocSvr.EDM.FileSystem.File.Move(visitpath & "\ClinicalVisit.pdf", visitpath & "\ClinicalVisit-Bkp1.pdf")
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\DetailLog.txt", Environment.NewLine & visitID + " Backup Done", True)
                        ''Creating new files tif, pdf, rtf
                        Decorator.DoWithTempFile(Sub(tmp)
                                                     CreateMultipageTIF(inputImages, tmp)
                                                     appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, visitpath & "\ClinicalVisit.tif")
                                                     TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                                     appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, visitpath & "\ClinicalVisit.rtf")
                                                     TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                                     appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, visitpath & "\ClinicalVisit.pdf")
                                                 End Sub)
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\DetailLog.txt", Environment.NewLine & visitID + " Process Ended", True)
                        IsChange = False
                        txtFileName.Text = ""
                    Else
                        My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & visitID + " No Such issue found", True)
                    End If
                End Using
            Catch ex As Exception
                ErrorList = ErrorList + "---" + visitID + " has Issue"
            End Try
            System.Net.ServicePointManager.SecurityProtocol = lastProtocol
            IsChange = False

        Catch ex As Exception
            ErrorList = ErrorList + "---" + visitID + " has Issue"
            My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & visitID + " has Issue", True)

        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = ExcelvisitString
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows

            If CheckLabResult(rows) Then

            Else

            End If

        Next

    End Sub
    Private Function CheckLabResult(ByVal strRTFFilePath As DataRow) As Boolean
        Dim visitID As String = strRTFFilePath.Item("SEQ_NUM").ToString
        Using TmpTextcontrol As New TXTextControl.ServerTextControl
            If Not TmpTextcontrol.Create Then
                Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
            End If

            TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)

            Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(strRTFFilePath.Item("PATH").ToString, "*.xml")
            If TmpTextcontrol.Text.Contains("Lab") AndAlso Not TmpTextcontrol.Text.Contains("Lab Results") Then
                My.Computer.FileSystem.WriteAllText("D:\Currently working\NewTemp\log.txt", Environment.NewLine & strRTFFilePath.Item("SEQ_NUM").ToString + "", True)
                For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields
                    If Not tmpfield.Name = "28" Then
                        Continue For
                    End If

                    Dim strdata As String = tmpfield.Text
                    Dim strolddata As String = ""
                    Dim StrCPTs As String = ""
                    For Each fileName As String In files
                        If Not fileName.Contains("28.xml") Then
                            Continue For
                        End If

                        Dim ds As New DataSet
                        Using ms As New MemoryStream(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(fileName))
                            ds.ReadXml(ms)
                        End Using


                        Dim dom As New HtmlDocument
                        dom.LoadHtml(ds.Tables("SECTION_ROW").Rows(0).Item("SECTION_HTML"))
                        Dim strsectionhtml As String = dom.DocumentNode.OuterHtml
                        Dim Nodes As HtmlNodeCollection = dom.DocumentNode.SelectNodes("//span")
                        Dim htmlTempElement As HtmlNode = dom.GetElementbyId("2838")

                        Dim insertSpan As String = "<span style='FONT-FAMILY: Arial; COLOR: blue; FONT-SIZE: 11pt; FONT-WEIGHT: bold'> Lab Results<br>" & "<span style='FONT-FAMILY: Arial; COLOR: blue; FONT-SIZE: 11pt; FONT-WEIGHT: bold'>" & strRTFFilePath.Item("TEST").ToString & " " & CType(strRTFFilePath.Item("ENTRY_DATE"), Date).ToShortDateString & "<br>" & "<span style= 'FONT-FAMILY: Arial; COLOR: blue; FONT-SIZE: 10pt; FONT-WEIGHT: normal'>" & strRTFFilePath.Item("TEST").ToString & "&nbsp;&nbsp;&nbsp;&nbsp;" & strRTFFilePath.Item("RESULT_VALUE").ToString & "<br><br></span></span>"
                        strsectionhtml = strsectionhtml.Replace(htmlTempElement.OuterHtml, htmlTempElement.OuterHtml & insertSpan) 'htmlTempElement.InnerHtml = htmlTempElement +  

                        Dim sectionSeqNum As String = ds.Tables("SECTION_ROW").Rows(0).Item("SECTION_SEQ_NUM").ToString
                        Dim eleSection As mshtml.HTMLSpanElement

                        Dim _HTMLDocument As mshtml.HTMLDocument = Nothing
                        Dim htmlDoc As mshtml.IHTMLDocument2 = CType(New mshtml.HTMLDocument, mshtml.IHTMLDocument2)
                        htmlDoc.write(strsectionhtml)
                        _HTMLDocument = New mshtml.HTMLDocument()
                        _HTMLDocument = CType(htmlDoc, mshtml.HTMLDocument)
                        eleSection = CType(_HTMLDocument.getElementById(sectionSeqNum), mshtml.HTMLSpanElement)

                        TmpTextcontrol.Selection.Start = tmpfield.Start - 1
                        TmpTextcontrol.Selection.Length = tmpfield.Length
                        Dim sectionHtml As String = RemoveTagsFromHTMl(eleSection.outerHTML.ToString)
                        Dim font As System.Drawing.Font = New System.Drawing.Font("Arial", 10.0!)
                        sectionHtml = "<SPAN" + " style='FONT-WEIGHT: normal; FONT-SIZE: " + font.SizeInPoints.ToString + "pt; FONT-FAMILY: " + font.Name + "'" + ">" + sectionHtml + "</SPAN>"
                        'TODO: Need to review, fonts are configureable in SequelMedEMR
                        TmpTextcontrol.Selection.Load(sectionHtml, TXTextControl.StringStreamType.HTMLFormat)
                        TmpTextcontrol.Selection.Start = tmpfield.Start - 1
                        TmpTextcontrol.Selection.Length = tmpfield.Length
                        Try
                            TmpTextcontrol.Selection.Save(sectionHtml, StringStreamType.HTMLFormat)
                            Exit For
                        Catch ex As Exception

                        End Try

                    Next
                    If tmpfield.ID = TmpTextcontrol.TextFields.Count - 4 Then
                        maxID = tmpfield.ID
                    End If
                    Exit For
                Next

            End If


            Dim PageCount As Integer = 0
            Dim inputImages As New ArrayList()

            For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                Dim image As New MemoryStream()
                ' get the image from TX Text Control's page rendering engine
                Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                ' save and add the image to the ArrayList
                mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                inputImages.Add(image)
                PageCount += 1
            Next
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\" & visitID & "xmls.zip", strRTFFilePath.Item("Path") & "\" & visitID & "xmls-Bkp.zip")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\ClinicalVisit.tif", strRTFFilePath.Item("Path") & "\ClinicalVisit-Bkp.tif")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath.Item("Path") & "\ClinicalVisit.rtf", strRTFFilePath.Item("Path") & "\ClinicalVisit-Bkp.rtf")

            Decorator.DoWithTempFile(Sub(tmp)
                                         CreateMultipageTIF(inputImages, tmp)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.tif")
                                         TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath.Item("PATH") & "\ClinicalVisit.rtf")
                                     End Sub)
        End Using
        Return True

    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim pram As OleDbParameter
        Dim dr As DataRow
        Dim olecon As OleDbConnection
        Dim olecomm As OleDbCommand
        Dim oleadpt As OleDbDataAdapter
        Dim ds As DataSet
        Try
            olecon = New OleDbConnection
            olecon.ConnectionString = PSANvisits
            olecomm = New OleDbCommand
            olecomm.CommandText =
               "Select * from [Visit$]"
            olecomm.Connection = olecon
            oleadpt = New OleDbDataAdapter(olecomm)
            ds = New DataSet
            olecon.Open()
            oleadpt.Fill(ds, "Sheet1")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            olecon.Close()
            olecon = Nothing
            olecomm = Nothing
            oleadpt = Nothing
            ds = Nothing
            dr = Nothing
            pram = Nothing
        End Try

        For Each rows As DataRow In ds.Tables(0).Rows
            Dim visitID As String = Path.GetFileName(rows.Item("TO_COPY").ToString)
            If Not appCtx.DocSvr.EDM.FileSystem.Directory.Exists(rows.Item("TO_COPY").ToString) Then
                Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(rows.Item("FROM_COPY").ToString, "*.xml", [option]:=SearchOption.AllDirectories)

                If Not files.Empty() Then
                    For Each fileName As String In files
                        Try
                            If Not fileName.Contains("PickList") AndAlso Not fileName.Contains("27.xml") AndAlso Not fileName.Contains("28.xml") Then
                                ''Copy files form past visit
                                appCtx.DocSvr.EDM.FileSystem.File.Copy(rows.Item("FROM_COPY") + "\" + Path.GetFileName(fileName), rows.Item("TO_COPY") + "\" + Path.GetFileName(fileName), True)
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                End If
            End If
            ''get flow sheet
            Dim strQuery As String
            strQuery = "select * From flow_sheet where CLINICAL_VISIT_SEQ_NUM = " + visitID
            Dim dsflowsheet As DataSet = Database.Select(strQuery)

            ''access and open each XML 
            FixXML(rows.Item("TO_COPY"), dsflowsheet)
        Next
        txtFileName.Text = ErrorList


    End Sub
    Private Function FixXML(ByVal VisitPath As String, ByVal Flowsheet As DataSet) As Boolean
        Try
            Dim visitID As String = Path.GetFileName(VisitPath)
            Dim lastProtocol As SecurityProtocolType = System.Net.ServicePointManager.SecurityProtocol
            System.Net.ServicePointManager.SecurityProtocol = CType(3072, System.Net.SecurityProtocolType)
            Dim files As String() = appCtx.DocSvr.EDM.FileSystem.Directory.GetFiles(VisitPath)
            If Not files.Empty() Then
                For Each fileName As String In files
                    Try

                        If Not fileName.Contains("PickList") AndAlso Not fileName.Contains("27.xml") AndAlso Not fileName.Contains("28.xml") AndAlso Not fileName.Contains("SentenceView.html") Then
                            If fileName.Contains("Backup") OrElse fileName.Contains("ICDCODES") Then
                                Continue For
                            End If
                            Dim dv As New DataView(Flowsheet.Tables(0))
                            dv.RowFilter = "META_SECTION_SEQ_NUM = " & Path.GetFileName(fileName).Replace(".xml", "")
                            dv.Sort = "META_QUESTION_GROUP_SEQ_NUM DESC"

                            'appCtx.DocSvr.EDM.FileSystem.File.Copy(VisitPath & "\" & Path.GetFileName(fileName), "D:\PSANTemp\" & Path.GetFileName(fileName))
                            appCtx.DocSvr.EDM.FileSystem.File.Copy(VisitPath & "\" & Path.GetFileName(fileName), VisitPath & "\" & Path.GetFileName(fileName) + "Backup")
                            Dim ds As New DataSet
                            Dim tempds As DataSet = Nothing
                            ds = appCtx.DocSvr.EDM.FileSystem.DataSet.ReadXml(fileName, tempds)
                            Dim IsChange As Boolean = False
                            For Each rowView As DataRowView In dv
                                Dim row As DataRow = rowView.Row
                                For Each xmlrow As DataRow In ds.Tables("QUESTION_ROW").Rows

                                    If String.IsNullOrEmpty(row.Item("QUESTION_SEQ_NUM").ToString) Then
                                        Continue For
                                    End If

                                    If row.Item("QUESTION_SEQ_NUM") = xmlrow.Item("QUESTION_SEQ_NUM") Then

                                        If row.Item("ANSWER_TYPE") = "META_QUESTION_BOOLEAN" Then
                                            xmlrow.Item("ANSWER_VALUE") = row.Item("ANSWER_BOOL_VALUE")
                                            xmlrow.Item("ANSWER_ENTERED_BY") = ""
                                            xmlrow.Item("ANSWER_DATE_TIME") = ""
                                            IsChange = True
                                        End If

                                        If row.Item("ANSWER_TYPE") = "META_QUESTION_NUMBER" Then
                                            xmlrow.Item("ANSWER_VALUE") = row.Item("ANSWER_INT_VALUE")
                                            xmlrow.Item("ANSWER_ENTERED_BY") = ""
                                            xmlrow.Item("ANSWER_DATE_TIME") = ""
                                            IsChange = True
                                        End If

                                        If String.IsNullOrEmpty(row.Item("ANSWER_TYPE")) OrElse row.Item("ANSWER_TYPE") = "META_QUESTION_STRING" OrElse row.Item("ANSWER_TYPE") = "META_QUESTION_DATE" OrElse row.Item("ANSWER_TYPE") = "META_QUESTION_DATE" OrElse row.Item("ANSWER_TYPE") = "META_QUESTION_DATETIME" OrElse row.Item("ANSWER_TYPE") = "META_QUESTION_IMAGE" OrElse row.Item("ANSWER_TYPE") = "META_QUESTION_IMAGE_LIST" Then
                                            xmlrow.Item("ANSWER_VALUE") = row.Item("ANSWER_STRING_VALUE")
                                            xmlrow.Item("ANSWER_ENTERED_BY") = ""
                                            xmlrow.Item("ANSWER_DATE_TIME") = ""
                                            IsChange = True
                                        End If

                                        If row.Item("ANSWER_TYPE") = "META_QUESTION_PICKLIST" Then
                                            Dim strQuery As String
                                            strQuery = "SELECT * FROM SEQUEL1.ANSWER_PICKLIST WHERE ANSWER_SEQ_NUM = " + row.Item("ANSWER_SEQ_NUM")
                                            Dim dsPicklistVal As DataSet = Database.Select(strQuery)
                                            xmlrow.Item("PICK_LIST_TEXT") = dsPicklistVal.Tables(0).Rows(0).Item("VALUE")
                                            xmlrow.Item("ANSWER_VALUE") = dsPicklistVal.Tables(0).Rows(0).Item("PICKLIST_ELEMENT_ID")
                                            xmlrow.Item("PICK_LIST_NUMERIC_VALUE") = dsPicklistVal.Tables(0).Rows(0).Item("NUMERIC")
                                            IsChange = True
                                        End If
                                    End If
                                Next

                            Next
                            If IsChange Then
                                ds.Tables("SECTION_ROW").Row.Item("SECTION_HTML") = ""
                            End If
                            Dim isWrite As Boolean = appCtx.DocSvr.EDM.FileSystem.DataSet.WriteXml(fileName, ds, XmlWriteMode.IgnoreSchema)
                        End If
                        System.Net.ServicePointManager.SecurityProtocol = lastProtocol
                    Catch ex As Exception

                    End Try
                Next
            End If
        Catch ex As Exception

        End Try
        Return True
    End Function

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If txtFileName.Text = "" Then
            Dim files As String() = Directory.GetFiles("C:\Users\mmasif\Downloads\Viits")

            If Not files.Empty() Then

                For Each fileName As String In files
                    If fileName.Contains(".pdf") Then
                        Continue For
                    End If
                    Dim inpFile As String = fileName
                    Dim outFile As String = fileName.Replace(".pdf", ".rtf")
                    Dim pdfLO As New PdfLoadOptions() With
                    {
                        .RasterizeVectorGraphics = False,
                        .DetectTables = False,
                        .PreserveEmbeddedFonts = False
                    }

                    Dim dc As DocumentCore = DocumentCore.Load(inpFile, pdfLO)
                    dc.Save(outFile)
                    ' Dim strRTFFilePath As String = "D:\PEHR DOS ChangerVB\PEHR DOS ChangerVB\bin\Debug\"

                    'File.Move(txtFileName.Text & "\ClinicalVisit.tif", txtFileName.Text & "\ClinicalVisit-Bkp.tif")
                    'Using TmpTextcontrol As New TXTextControl.ServerTextControl
                    '    If Not TmpTextcontrol.Create Then
                    '        Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
                    '    End If

                    '    TmpTextcontrol.Load((inpFile), TXTextControl.StreamType.RichTextFormat)

                    '    Dim PageCount As Integer = 0
                    '    Dim inputImages As New ArrayList()
                    '    For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                    '        Dim image As New MemoryStream()
                    '        ' get the image from TX Text Control's page rendering engine
                    '        Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                    '        ' save and add the image to the ArrayList
                    '        mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                    '        inputImages.Add(image)
                    '        PageCount += 1
                    '    Next
                    '    Decorator.DoWithTempFile(Sub(tmp)
                    '                                 CreateMultipageTIF(inputImages, tmp)
                    '                                 File.Move(tmp, fileName.Replace(".rtf", ".tif"))
                    '                             End Sub)
                    'End Using

                    'System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
                Next
            End If

        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim strRTFFilePath As String = txtFileName.Text
        ' File.Move(txtFileName.Text & "\ClinicalVisit.tif", txtFileName.Text & "\ClinicalVisit-Bkp.tif")
        Using TmpTextcontrol As New TXTextControl.ServerTextControl
            If Not TmpTextcontrol.Create Then
                Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
            End If

            'TmpTextcontrol.Load((strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
            TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)

            Dim PageCount As Integer = 0
            Dim inputImages As New ArrayList()
            For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                Dim image As New MemoryStream()
                ' get the image from TX Text Control's page rendering engine
                Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                ' save and add the image to the ArrayList
                mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                inputImages.Add(image)
                PageCount += 1
            Next

            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.tif", strRTFFilePath & "\ClinicalVisit-Bkp2.tif")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.rtf", strRTFFilePath & "\ClinicalVisit-Bkp2.rtf")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.pdf", strRTFFilePath & "\ClinicalVisit-Bkp2.pdf")
            ''Creating new files tif, pdf, rtf
            Decorator.DoWithTempFile(Sub(tmp)
                                         CreateMultipageTIF(inputImages, tmp)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.tif")
                                         TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.rtf")
                                         TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.pdf")
                                     End Sub)

        End Using
        txtFileName.Text = ""
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim strRTFFilePath As String = txtFileName.Text
        ' File.Move(txtFileName.Text & "\ClinicalVisit.tif", txtFileName.Text & "\ClinicalVisit-Bkp.tif")
        Using TmpTextcontrol As New TXTextControl.ServerTextControl
            If Not TmpTextcontrol.Create Then
                Throw New Exception("Unable to initialize TXTextControl.ServerTextControl")
            End If

            'TmpTextcontrol.Load((strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)
            TmpTextcontrol.Load(appCtx.DocSvr.EDM.FileSystem.File.ReadAllBytes(strRTFFilePath & "\ClinicalVisit.rtf"), TXTextControl.StreamType.RichTextFormat)

            For Each tmpfield As TXTextControl.TextField In TmpTextcontrol.TextFields

                If tmpfield.Name.Contains("CLINICAL_VISIT_DATE") Then
                    tmpfield.Text = txtNewDOS.Text
                End If

            Next

            Dim PageCount As Integer = 0
            Dim inputImages As New ArrayList()
            For Each page As TXTextControl.Page In TmpTextcontrol.GetPages()
                Dim image As New MemoryStream()
                ' get the image from TX Text Control's page rendering engine
                Dim mf As Bitmap = page.GetImage(100, TXTextControl.Page.PageContent.All)
                ' save and add the image to the ArrayList
                mf.Save(image, System.Drawing.Imaging.ImageFormat.Tiff)
                inputImages.Add(image)
                PageCount += 1
            Next

            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.tif", strRTFFilePath & "\ClinicalVisit-Bkp2.tif")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.rtf", strRTFFilePath & "\ClinicalVisit-Bkp2.rtf")
            appCtx.DocSvr.EDM.FileSystem.File.Move(strRTFFilePath & "\ClinicalVisit.pdf", strRTFFilePath & "\ClinicalVisit-Bkp2.pdf")
            ''Creating new files tif, pdf, rtf
            Decorator.DoWithTempFile(Sub(tmp)
                                         CreateMultipageTIF(inputImages, tmp)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.tif")
                                         TmpTextcontrol.Save(tmp, TXTextControl.StreamType.RichTextFormat)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.rtf")
                                         TmpTextcontrol.Save(tmp, TXTextControl.StreamType.AdobePDF)
                                         appCtx.DocSvr.EDM.FileSystem.File.Move(tmp, strRTFFilePath & "\ClinicalVisit.pdf")
                                     End Sub)

        End Using
        txtFileName.Text = ""
    End Sub
End Class
