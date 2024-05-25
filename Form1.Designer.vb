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
        ListBox1 = New ListBox()
        ListBox2 = New ListBox()
        ListBox3 = New ListBox()
        ListBox4 = New ListBox()
        Button1 = New Button()
        Button2 = New Button()
        Button3 = New Button()
        SuspendLayout()
        ' 
        ' ListBox1
        ' 
        ListBox1.FormattingEnabled = True
        ListBox1.HorizontalScrollbar = True
        ListBox1.ItemHeight = 17
        ListBox1.Location = New Point(12, 42)
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(465, 548)
        ListBox1.TabIndex = 0
        ' 
        ' ListBox2
        ' 
        ListBox2.FormattingEnabled = True
        ListBox2.HorizontalScrollbar = True
        ListBox2.ItemHeight = 17
        ListBox2.Location = New Point(483, 42)
        ListBox2.Name = "ListBox2"
        ListBox2.Size = New Size(465, 548)
        ListBox2.TabIndex = 1
        ' 
        ' ListBox3
        ' 
        ListBox3.FormattingEnabled = True
        ListBox3.HorizontalScrollbar = True
        ListBox3.ItemHeight = 17
        ListBox3.Location = New Point(954, 42)
        ListBox3.Name = "ListBox3"
        ListBox3.Size = New Size(465, 548)
        ListBox3.TabIndex = 2
        ' 
        ' ListBox4
        ' 
        ListBox4.FormattingEnabled = True
        ListBox4.HorizontalScrollbar = True
        ListBox4.ItemHeight = 17
        ListBox4.Location = New Point(12, 596)
        ListBox4.Name = "ListBox4"
        ListBox4.Size = New Size(1407, 106)
        ListBox4.TabIndex = 3
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(12, 12)
        Button1.Name = "Button1"
        Button1.Size = New Size(90, 24)
        Button1.TabIndex = 4
        Button1.Text = "打开文件夹"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Enabled = False
        Button2.Location = New Point(108, 12)
        Button2.Name = "Button2"
        Button2.Size = New Size(90, 24)
        Button2.TabIndex = 5
        Button2.Text = "获取文件列表"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' Button3
        ' 
        Button3.Enabled = False
        Button3.Location = New Point(204, 12)
        Button3.Name = "Button3"
        Button3.Size = New Size(90, 24)
        Button3.TabIndex = 6
        Button3.Text = "文件夹对比"
        Button3.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1431, 714)
        Controls.Add(Button3)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(ListBox4)
        Controls.Add(ListBox3)
        Controls.Add(ListBox2)
        Controls.Add(ListBox1)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "文件夹对比V1.0.0"
        ResumeLayout(False)
    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents ListBox2 As ListBox
    Friend WithEvents ListBox3 As ListBox
    Friend WithEvents ListBox4 As ListBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button

End Class
