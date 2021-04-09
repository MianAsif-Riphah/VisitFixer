<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnFileOpen = New System.Windows.Forms.Button()
        Me.btnChangeDOS = New System.Windows.Forms.Button()
        Me.txtFileName = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnGetSectionData = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnGetCurrptFiles = New System.Windows.Forms.Button()
        Me.btnGetUniqSec = New System.Windows.Forms.Button()
        Me.btnCreateTiff = New System.Windows.Forms.Button()
        Me.btnGetPlanIssue = New System.Windows.Forms.Button()
        Me.btnHeaderSig = New System.Windows.Forms.Button()
        Me.btnPastVisitFileCopy = New System.Windows.Forms.Button()
        Me.btnSWHeaderimg = New System.Windows.Forms.Button()
        Me.btnSDFAPastVisitFix = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.txtNewDOS = New System.Windows.Forms.TextBox()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnFileOpen
        '
        Me.btnFileOpen.Location = New System.Drawing.Point(327, 73)
        Me.btnFileOpen.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFileOpen.Name = "btnFileOpen"
        Me.btnFileOpen.Size = New System.Drawing.Size(36, 25)
        Me.btnFileOpen.TabIndex = 0
        Me.btnFileOpen.Text = "..."
        Me.btnFileOpen.UseVisualStyleBackColor = True
        '
        'btnChangeDOS
        '
        Me.btnChangeDOS.Location = New System.Drawing.Point(16, 105)
        Me.btnChangeDOS.Margin = New System.Windows.Forms.Padding(4)
        Me.btnChangeDOS.Name = "btnChangeDOS"
        Me.btnChangeDOS.Size = New System.Drawing.Size(121, 28)
        Me.btnChangeDOS.TabIndex = 1
        Me.btnChangeDOS.Text = "Change DOS"
        Me.btnChangeDOS.UseVisualStyleBackColor = True
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(16, 73)
        Me.txtFileName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(301, 22)
        Me.txtFileName.TabIndex = 2
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btnGetSectionData
        '
        Me.btnGetSectionData.Location = New System.Drawing.Point(145, 105)
        Me.btnGetSectionData.Margin = New System.Windows.Forms.Padding(4)
        Me.btnGetSectionData.Name = "btnGetSectionData"
        Me.btnGetSectionData.Size = New System.Drawing.Size(173, 28)
        Me.btnGetSectionData.TabIndex = 3
        Me.btnGetSectionData.Text = "Get Data from section"
        Me.btnGetSectionData.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(163, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Result"
        '
        'BtnGetCurrptFiles
        '
        Me.BtnGetCurrptFiles.Location = New System.Drawing.Point(16, 140)
        Me.BtnGetCurrptFiles.Margin = New System.Windows.Forms.Padding(4)
        Me.BtnGetCurrptFiles.Name = "BtnGetCurrptFiles"
        Me.BtnGetCurrptFiles.Size = New System.Drawing.Size(121, 28)
        Me.BtnGetCurrptFiles.TabIndex = 5
        Me.BtnGetCurrptFiles.Text = "Get Currpt Files"
        Me.BtnGetCurrptFiles.UseVisualStyleBackColor = True
        '
        'btnGetUniqSec
        '
        Me.btnGetUniqSec.Location = New System.Drawing.Point(145, 140)
        Me.btnGetUniqSec.Margin = New System.Windows.Forms.Padding(4)
        Me.btnGetUniqSec.Name = "btnGetUniqSec"
        Me.btnGetUniqSec.Size = New System.Drawing.Size(173, 28)
        Me.btnGetUniqSec.TabIndex = 6
        Me.btnGetUniqSec.Text = "Create Sentanceview"
        Me.btnGetUniqSec.UseVisualStyleBackColor = True
        '
        'btnCreateTiff
        '
        Me.btnCreateTiff.Location = New System.Drawing.Point(16, 176)
        Me.btnCreateTiff.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCreateTiff.Name = "btnCreateTiff"
        Me.btnCreateTiff.Size = New System.Drawing.Size(121, 28)
        Me.btnCreateTiff.TabIndex = 7
        Me.btnCreateTiff.Text = "Create tiff"
        Me.btnCreateTiff.UseVisualStyleBackColor = True
        '
        'btnGetPlanIssue
        '
        Me.btnGetPlanIssue.Location = New System.Drawing.Point(145, 176)
        Me.btnGetPlanIssue.Margin = New System.Windows.Forms.Padding(4)
        Me.btnGetPlanIssue.Name = "btnGetPlanIssue"
        Me.btnGetPlanIssue.Size = New System.Drawing.Size(173, 28)
        Me.btnGetPlanIssue.TabIndex = 8
        Me.btnGetPlanIssue.Text = "Get List Of Plan"
        Me.btnGetPlanIssue.UseVisualStyleBackColor = True
        '
        'btnHeaderSig
        '
        Me.btnHeaderSig.Location = New System.Drawing.Point(16, 212)
        Me.btnHeaderSig.Margin = New System.Windows.Forms.Padding(4)
        Me.btnHeaderSig.Name = "btnHeaderSig"
        Me.btnHeaderSig.Size = New System.Drawing.Size(303, 28)
        Me.btnHeaderSig.TabIndex = 9
        Me.btnHeaderSig.Text = "Place Header and Sig"
        Me.btnHeaderSig.UseVisualStyleBackColor = True
        '
        'btnPastVisitFileCopy
        '
        Me.btnPastVisitFileCopy.Location = New System.Drawing.Point(16, 283)
        Me.btnPastVisitFileCopy.Margin = New System.Windows.Forms.Padding(4)
        Me.btnPastVisitFileCopy.Name = "btnPastVisitFileCopy"
        Me.btnPastVisitFileCopy.Size = New System.Drawing.Size(303, 28)
        Me.btnPastVisitFileCopy.TabIndex = 10
        Me.btnPastVisitFileCopy.Text = "Past Visit File Copy"
        Me.btnPastVisitFileCopy.UseVisualStyleBackColor = True
        '
        'btnSWHeaderimg
        '
        Me.btnSWHeaderimg.Location = New System.Drawing.Point(16, 247)
        Me.btnSWHeaderimg.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSWHeaderimg.Name = "btnSWHeaderimg"
        Me.btnSWHeaderimg.Size = New System.Drawing.Size(303, 28)
        Me.btnSWHeaderimg.TabIndex = 11
        Me.btnSWHeaderimg.Text = "SW Place Header and Sig"
        Me.btnSWHeaderimg.UseVisualStyleBackColor = True
        '
        'btnSDFAPastVisitFix
        '
        Me.btnSDFAPastVisitFix.Location = New System.Drawing.Point(16, 319)
        Me.btnSDFAPastVisitFix.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSDFAPastVisitFix.Name = "btnSDFAPastVisitFix"
        Me.btnSDFAPastVisitFix.Size = New System.Drawing.Size(303, 28)
        Me.btnSDFAPastVisitFix.TabIndex = 12
        Me.btnSDFAPastVisitFix.Text = "SDFA Past visit fix"
        Me.btnSDFAPastVisitFix.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(16, 354)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(303, 28)
        Me.Button1.TabIndex = 13
        Me.Button1.Text = "Signture + Section Data fixing"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Location = New System.Drawing.Point(327, 118)
        Me.WebBrowser1.Margin = New System.Windows.Forms.Padding(4)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(27, 25)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(81, 50)
        Me.WebBrowser1.TabIndex = 14
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(16, 390)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(303, 28)
        Me.Button2.TabIndex = 15
        Me.Button2.Text = "Header Data Replace TIH"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(16, 426)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(303, 28)
        Me.Button3.TabIndex = 16
        Me.Button3.Text = "SignOff line Change"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(399, 426)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(241, 28)
        Me.Button4.TabIndex = 17
        Me.Button4.Text = "Excel Missing Lab result"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(399, 354)
        Me.Button5.Margin = New System.Windows.Forms.Padding(4)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(241, 28)
        Me.Button5.TabIndex = 18
        Me.Button5.Text = "PSAN Visit FIx"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(399, 319)
        Me.Button6.Margin = New System.Windows.Forms.Padding(4)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(241, 28)
        Me.Button6.TabIndex = 19
        Me.Button6.Text = "PSAN Visit FIx"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(399, 267)
        Me.Button7.Margin = New System.Windows.Forms.Padding(4)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(241, 28)
        Me.Button7.TabIndex = 20
        Me.Button7.Text = "Create VisitFiles"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'txtNewDOS
        '
        Me.txtNewDOS.Location = New System.Drawing.Point(456, 182)
        Me.txtNewDOS.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNewDOS.Name = "txtNewDOS"
        Me.txtNewDOS.Size = New System.Drawing.Size(98, 22)
        Me.txtNewDOS.TabIndex = 21
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(399, 231)
        Me.Button8.Margin = New System.Windows.Forms.Padding(4)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(241, 28)
        Me.Button8.TabIndex = 22
        Me.Button8.Text = "Update DOS"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(656, 459)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.txtNewDOS)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.WebBrowser1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnSDFAPastVisitFix)
        Me.Controls.Add(Me.btnSWHeaderimg)
        Me.Controls.Add(Me.btnPastVisitFileCopy)
        Me.Controls.Add(Me.btnHeaderSig)
        Me.Controls.Add(Me.btnGetPlanIssue)
        Me.Controls.Add(Me.btnCreateTiff)
        Me.Controls.Add(Me.btnGetUniqSec)
        Me.Controls.Add(Me.BtnGetCurrptFiles)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnGetSectionData)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.btnChangeDOS)
        Me.Controls.Add(Me.btnFileOpen)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnFileOpen As Button
    Friend WithEvents btnChangeDOS As Button
    Friend WithEvents txtFileName As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnGetSectionData As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents BtnGetCurrptFiles As Button
    Friend WithEvents btnGetUniqSec As Button
    Friend WithEvents btnCreateTiff As Button
    Friend WithEvents btnGetPlanIssue As Button
    Friend WithEvents btnHeaderSig As Button
    Friend WithEvents btnPastVisitFileCopy As Button
    Friend WithEvents btnSWHeaderimg As Button
    Friend WithEvents btnSDFAPastVisitFix As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents WebBrowser1 As WebBrowser
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents txtNewDOS As TextBox
    Friend WithEvents Button8 As Button
End Class
