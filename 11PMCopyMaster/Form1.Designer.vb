<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        btnTest = New Button()
        lbTestResults = New ListBox()
        Label1 = New Label()
        SuspendLayout()
        ' 
        ' btnTest
        ' 
        btnTest.Location = New Point(3, 12)
        btnTest.Name = "btnTest"
        btnTest.Size = New Size(150, 58)
        btnTest.TabIndex = 0
        btnTest.Text = "Test"
        btnTest.UseVisualStyleBackColor = True
        ' 
        ' lbTestResults
        ' 
        lbTestResults.FormattingEnabled = True
        lbTestResults.ItemHeight = 25
        lbTestResults.Location = New Point(180, 34)
        lbTestResults.Name = "lbTestResults"
        lbTestResults.Size = New Size(366, 404)
        lbTestResults.TabIndex = 1
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(187, 3)
        Label1.Name = "Label1"
        Label1.Size = New Size(102, 25)
        Label1.TabIndex = 2
        Label1.Text = "Test Results"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(10F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(566, 450)
        Controls.Add(Label1)
        Controls.Add(lbTestResults)
        Controls.Add(btnTest)
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "11pm Copy Master"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnTest As Button
    Friend WithEvents lbTestResults As ListBox
    Friend WithEvents Label1 As Label

End Class
