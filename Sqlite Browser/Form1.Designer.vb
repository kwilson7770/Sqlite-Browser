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
        Me.components = New System.ComponentModel.Container()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ComboBoxTables = New System.Windows.Forms.ComboBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ButtonSave = New System.Windows.Forms.Button()
        Me.ButtonLoadDB = New System.Windows.Forms.Button()
        Me.TextBoxDBPath = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ListBoxDebugger = New System.Windows.Forms.ListBox()
        Me.TimerDebugger = New System.Windows.Forms.Timer(Me.components)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 64)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(688, 297)
        Me.DataGridView1.TabIndex = 3
        '
        'ComboBoxTables
        '
        Me.ComboBoxTables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBoxTables.FormattingEnabled = True
        Me.ComboBoxTables.Location = New System.Drawing.Point(450, 36)
        Me.ComboBoxTables.Name = "ComboBoxTables"
        Me.ComboBoxTables.Size = New System.Drawing.Size(126, 21)
        Me.ComboBoxTables.TabIndex = 4
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ButtonSave)
        Me.Panel1.Controls.Add(Me.ButtonLoadDB)
        Me.Panel1.Controls.Add(Me.TextBoxDBPath)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ComboBoxTables)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(688, 64)
        Me.Panel1.TabIndex = 5
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(582, 34)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(94, 23)
        Me.ButtonSave.TabIndex = 9
        Me.ButtonSave.Text = "Save Changes"
        Me.ButtonSave.UseVisualStyleBackColor = True
        '
        'ButtonLoadDB
        '
        Me.ButtonLoadDB.Location = New System.Drawing.Point(353, 36)
        Me.ButtonLoadDB.Name = "ButtonLoadDB"
        Me.ButtonLoadDB.Size = New System.Drawing.Size(91, 23)
        Me.ButtonLoadDB.TabIndex = 8
        Me.ButtonLoadDB.Text = "Load Database"
        Me.ButtonLoadDB.UseVisualStyleBackColor = True
        '
        'TextBoxDBPath
        '
        Me.TextBoxDBPath.Location = New System.Drawing.Point(12, 38)
        Me.TextBoxDBPath.Name = "TextBoxDBPath"
        Me.TextBoxDBPath.Size = New System.Drawing.Size(335, 20)
        Me.TextBoxDBPath.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Database Path"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(447, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Tables"
        '
        'ListBoxDebugger
        '
        Me.ListBoxDebugger.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ListBoxDebugger.FormattingEnabled = True
        Me.ListBoxDebugger.HorizontalScrollbar = True
        Me.ListBoxDebugger.Location = New System.Drawing.Point(0, 361)
        Me.ListBoxDebugger.Name = "ListBoxDebugger"
        Me.ListBoxDebugger.Size = New System.Drawing.Size(688, 69)
        Me.ListBoxDebugger.TabIndex = 6
        Me.ListBoxDebugger.Visible = False
        '
        'TimerDebugger
        '
        Me.TimerDebugger.Interval = 500
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(688, 430)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ListBoxDebugger)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Sqlite Browser"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ComboBoxTables As ComboBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents TextBoxDBPath As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents ButtonLoadDB As Button
    Friend WithEvents ListBoxDebugger As ListBox
    Friend WithEvents TimerDebugger As Timer
    Friend WithEvents ButtonSave As Button
End Class
