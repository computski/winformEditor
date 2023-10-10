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
        DataGridView1 = New DataGridView()
        Button1 = New Button()
        Button2 = New Button()
        TextBox1 = New TextBox()
        Column1 = New DataGridViewComboBoxColumn()
        FirstName = New DataGridViewTextBoxColumn()
        IDplayer = New DataGridViewTextBoxColumn()
        CType(DataGridView1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' DataGridView1
        ' 
        DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridView1.Columns.AddRange(New DataGridViewColumn() {Column1, FirstName, IDplayer})
        DataGridView1.Location = New Point(78, 65)
        DataGridView1.Name = "DataGridView1"
        DataGridView1.RowTemplate.Height = 25
        DataGridView1.Size = New Size(579, 156)
        DataGridView1.TabIndex = 0
        ' 
        ' Button1
        ' 
        Button1.Location = New Point(88, 271)
        Button1.Name = "Button1"
        Button1.Size = New Size(75, 23)
        Button1.TabIndex = 1
        Button1.Text = "refresh"
        Button1.UseVisualStyleBackColor = True
        ' 
        ' Button2
        ' 
        Button2.Location = New Point(228, 271)
        Button2.Name = "Button2"
        Button2.Size = New Size(75, 23)
        Button2.TabIndex = 2
        Button2.Text = "Update"
        Button2.UseVisualStyleBackColor = True
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(79, 347)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New Size(284, 23)
        TextBox1.TabIndex = 3
        ' 
        ' Column1
        ' 
        Column1.HeaderText = "Column1"
        Column1.Name = "Column1"
        ' 
        ' FirstName
        ' 
        FirstName.DataPropertyName = "FirstName"
        FirstName.HeaderText = "First Name"
        FirstName.Name = "FirstName"
        ' 
        ' IDplayer
        ' 
        IDplayer.DataPropertyName = "IDplayer"
        IDplayer.HeaderText = "IDplayer"
        IDplayer.Name = "IDplayer"
        IDplayer.ReadOnly = True
        IDplayer.Visible = False
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(800, 450)
        Controls.Add(TextBox1)
        Controls.Add(Button2)
        Controls.Add(Button1)
        Controls.Add(DataGridView1)
        Name = "Form1"
        Text = "Form1"
        CType(DataGridView1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Column1 As DataGridViewComboBoxColumn
    Friend WithEvents FirstName As DataGridViewTextBoxColumn
    Friend WithEvents IDplayer As DataGridViewTextBoxColumn
End Class
