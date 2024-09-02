<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Clavier
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Lettre_I = New System.Windows.Forms.Label()
        Me.Lettre_U = New System.Windows.Forms.Label()
        Me.Lettre_Y = New System.Windows.Forms.Label()
        Me.Lettre_R = New System.Windows.Forms.Label()
        Me.Lettre_E = New System.Windows.Forms.Label()
        Me.Lettre_L = New System.Windows.Forms.Label()
        Me.Lettre_K = New System.Windows.Forms.Label()
        Me.Lettre_J = New System.Windows.Forms.Label()
        Me.Lettre_H = New System.Windows.Forms.Label()
        Me.Lettre_G = New System.Windows.Forms.Label()
        Me.Lettre_F = New System.Windows.Forms.Label()
        Me.Lettre_D = New System.Windows.Forms.Label()
        Me.Lettre_S = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.VolumeEcoute = New System.Windows.Forms.TrackBar()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ListPRG = New System.Windows.Forms.ComboBox()
        Me.CanalEcoute = New System.Windows.Forms.NumericUpDown()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.VéloEcoute = New System.Windows.Forms.NumericUpDown()
        Me.OctaveEcoute = New System.Windows.Forms.NumericUpDown()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ValCtrl = New System.Windows.Forms.TrackBar()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.VolumeEcoute, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CanalEcoute, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VéloEcoute, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OctaveEcoute, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        CType(Me.ValCtrl, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Khaki
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Lettre_I)
        Me.Panel1.Controls.Add(Me.Lettre_U)
        Me.Panel1.Controls.Add(Me.Lettre_Y)
        Me.Panel1.Controls.Add(Me.Lettre_R)
        Me.Panel1.Controls.Add(Me.Lettre_E)
        Me.Panel1.Controls.Add(Me.Lettre_L)
        Me.Panel1.Controls.Add(Me.Lettre_K)
        Me.Panel1.Controls.Add(Me.Lettre_J)
        Me.Panel1.Controls.Add(Me.Lettre_H)
        Me.Panel1.Controls.Add(Me.Lettre_G)
        Me.Panel1.Controls.Add(Me.Lettre_F)
        Me.Panel1.Controls.Add(Me.Lettre_D)
        Me.Panel1.Controls.Add(Me.Lettre_S)
        Me.Panel1.Location = New System.Drawing.Point(7, 217)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(337, 74)
        Me.Panel1.TabIndex = 0
        '
        'Lettre_I
        '
        Me.Lettre_I.BackColor = System.Drawing.Color.White
        Me.Lettre_I.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_I.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_I.Location = New System.Drawing.Point(226, 12)
        Me.Lettre_I.Name = "Lettre_I"
        Me.Lettre_I.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_I.TabIndex = 12
        Me.Lettre_I.Text = "I"
        Me.Lettre_I.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_U
        '
        Me.Lettre_U.BackColor = System.Drawing.Color.White
        Me.Lettre_U.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_U.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_U.Location = New System.Drawing.Point(186, 12)
        Me.Lettre_U.Name = "Lettre_U"
        Me.Lettre_U.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_U.TabIndex = 11
        Me.Lettre_U.Text = "U"
        Me.Lettre_U.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_Y
        '
        Me.Lettre_Y.BackColor = System.Drawing.Color.White
        Me.Lettre_Y.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_Y.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_Y.Location = New System.Drawing.Point(147, 12)
        Me.Lettre_Y.Name = "Lettre_Y"
        Me.Lettre_Y.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_Y.TabIndex = 10
        Me.Lettre_Y.Text = "Y"
        Me.Lettre_Y.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_R
        '
        Me.Lettre_R.BackColor = System.Drawing.Color.White
        Me.Lettre_R.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_R.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_R.Location = New System.Drawing.Point(69, 12)
        Me.Lettre_R.Name = "Lettre_R"
        Me.Lettre_R.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_R.TabIndex = 9
        Me.Lettre_R.Text = "R"
        Me.Lettre_R.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_E
        '
        Me.Lettre_E.BackColor = System.Drawing.Color.White
        Me.Lettre_E.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_E.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_E.Location = New System.Drawing.Point(29, 12)
        Me.Lettre_E.Name = "Lettre_E"
        Me.Lettre_E.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_E.TabIndex = 8
        Me.Lettre_E.Text = "E"
        Me.Lettre_E.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_L
        '
        Me.Lettre_L.BackColor = System.Drawing.Color.White
        Me.Lettre_L.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_L.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_L.Location = New System.Drawing.Point(288, 36)
        Me.Lettre_L.Name = "Lettre_L"
        Me.Lettre_L.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_L.TabIndex = 7
        Me.Lettre_L.Text = "L"
        Me.Lettre_L.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_K
        '
        Me.Lettre_K.BackColor = System.Drawing.Color.White
        Me.Lettre_K.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_K.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_K.Location = New System.Drawing.Point(248, 36)
        Me.Lettre_K.Name = "Lettre_K"
        Me.Lettre_K.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_K.TabIndex = 6
        Me.Lettre_K.Text = "K"
        Me.Lettre_K.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_J
        '
        Me.Lettre_J.BackColor = System.Drawing.Color.White
        Me.Lettre_J.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_J.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_J.Location = New System.Drawing.Point(208, 36)
        Me.Lettre_J.Name = "Lettre_J"
        Me.Lettre_J.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_J.TabIndex = 5
        Me.Lettre_J.Text = "J"
        Me.Lettre_J.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_H
        '
        Me.Lettre_H.BackColor = System.Drawing.Color.White
        Me.Lettre_H.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_H.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_H.Location = New System.Drawing.Point(168, 36)
        Me.Lettre_H.Name = "Lettre_H"
        Me.Lettre_H.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_H.TabIndex = 4
        Me.Lettre_H.Text = "H"
        Me.Lettre_H.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_G
        '
        Me.Lettre_G.BackColor = System.Drawing.Color.White
        Me.Lettre_G.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_G.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_G.Location = New System.Drawing.Point(128, 36)
        Me.Lettre_G.Name = "Lettre_G"
        Me.Lettre_G.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_G.TabIndex = 3
        Me.Lettre_G.Text = "G"
        Me.Lettre_G.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_F
        '
        Me.Lettre_F.BackColor = System.Drawing.Color.White
        Me.Lettre_F.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_F.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_F.Location = New System.Drawing.Point(88, 36)
        Me.Lettre_F.Name = "Lettre_F"
        Me.Lettre_F.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_F.TabIndex = 2
        Me.Lettre_F.Text = "F"
        Me.Lettre_F.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_D
        '
        Me.Lettre_D.BackColor = System.Drawing.Color.White
        Me.Lettre_D.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_D.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_D.Location = New System.Drawing.Point(48, 36)
        Me.Lettre_D.Name = "Lettre_D"
        Me.Lettre_D.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_D.TabIndex = 1
        Me.Lettre_D.Text = "D"
        Me.Lettre_D.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Lettre_S
        '
        Me.Lettre_S.BackColor = System.Drawing.Color.White
        Me.Lettre_S.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Lettre_S.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lettre_S.Location = New System.Drawing.Point(8, 36)
        Me.Lettre_S.Name = "Lettre_S"
        Me.Lettre_S.Size = New System.Drawing.Size(34, 23)
        Me.Lettre_S.TabIndex = 0
        Me.Lettre_S.Text = "S"
        Me.Lettre_S.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.AntiqueWhite
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.CheckBox1)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.VolumeEcoute)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.ListPRG)
        Me.Panel2.Controls.Add(Me.CanalEcoute)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.VéloEcoute)
        Me.Panel2.Controls.Add(Me.OctaveEcoute)
        Me.Panel2.Controls.Add(Me.Label18)
        Me.Panel2.Location = New System.Drawing.Point(8, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(336, 124)
        Me.Panel2.TabIndex = 1
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(206, 46)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(118, 17)
        Me.CheckBox1.TabIndex = 208
        Me.CheckBox1.Text = "Toujours au dessus"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(7, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 207
        Me.Label3.Text = "Volume"
        '
        'VolumeEcoute
        '
        Me.VolumeEcoute.Location = New System.Drawing.Point(4, 69)
        Me.VolumeEcoute.Maximum = 127
        Me.VolumeEcoute.Name = "VolumeEcoute"
        Me.VolumeEcoute.Size = New System.Drawing.Size(317, 45)
        Me.VolumeEcoute.TabIndex = 206
        Me.VolumeEcoute.Value = 95
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(156, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(30, 13)
        Me.Label2.TabIndex = 205
        Me.Label2.Text = "PRG"
        '
        'ListPRG
        '
        Me.ListPRG.FormattingEnabled = True
        Me.ListPRG.Location = New System.Drawing.Point(159, 9)
        Me.ListPRG.Name = "ListPRG"
        Me.ListPRG.Size = New System.Drawing.Size(162, 21)
        Me.ListPRG.TabIndex = 204
        '
        'CanalEcoute
        '
        Me.CanalEcoute.Location = New System.Drawing.Point(104, 9)
        Me.CanalEcoute.Maximum = New Decimal(New Integer() {16, 0, 0, 0})
        Me.CanalEcoute.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.CanalEcoute.Name = "CanalEcoute"
        Me.CanalEcoute.Size = New System.Drawing.Size(39, 20)
        Me.CanalEcoute.TabIndex = 202
        Me.CanalEcoute.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(104, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 13)
        Me.Label1.TabIndex = 203
        Me.Label1.Text = "Canal"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(58, 34)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(45, 13)
        Me.Label14.TabIndex = 201
        Me.Label14.Text = "Vélocité"
        '
        'VéloEcoute
        '
        Me.VéloEcoute.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VéloEcoute.Location = New System.Drawing.Point(56, 9)
        Me.VéloEcoute.Maximum = New Decimal(New Integer() {127, 0, 0, 0})
        Me.VéloEcoute.Name = "VéloEcoute"
        Me.VéloEcoute.Size = New System.Drawing.Size(40, 20)
        Me.VéloEcoute.TabIndex = 200
        Me.VéloEcoute.Value = New Decimal(New Integer() {90, 0, 0, 0})
        '
        'OctaveEcoute
        '
        Me.OctaveEcoute.Location = New System.Drawing.Point(8, 9)
        Me.OctaveEcoute.Maximum = New Decimal(New Integer() {6, 0, 0, 0})
        Me.OctaveEcoute.Minimum = New Decimal(New Integer() {2, 0, 0, -2147483648})
        Me.OctaveEcoute.Name = "OctaveEcoute"
        Me.OctaveEcoute.Size = New System.Drawing.Size(40, 20)
        Me.OctaveEcoute.TabIndex = 198
        Me.OctaveEcoute.Value = New Decimal(New Integer() {3, 0, 0, 0})
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(7, 34)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(42, 13)
        Me.Label18.TabIndex = 199
        Me.Label18.Text = "Octave"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Ivory
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.ComboBox1)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.ValCtrl)
        Me.Panel3.Location = New System.Drawing.Point(7, 133)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(337, 78)
        Me.Panel3.TabIndex = 2
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Calibri", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(236, 26)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(93, 21)
        Me.ComboBox1.TabIndex = 209
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(60, 13)
        Me.Label4.TabIndex = 208
        Me.Label4.Text = "MIDI Learn"
        '
        'ValCtrl
        '
        Me.ValCtrl.BackColor = System.Drawing.Color.Ivory
        Me.ValCtrl.Location = New System.Drawing.Point(3, 20)
        Me.ValCtrl.Maximum = 127
        Me.ValCtrl.Name = "ValCtrl"
        Me.ValCtrl.Size = New System.Drawing.Size(227, 45)
        Me.ValCtrl.TabIndex = 207
        Me.ValCtrl.Value = 95
        '
        'Clavier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(350, 299)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Clavier"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Clavier"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.VolumeEcoute, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CanalEcoute, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VéloEcoute, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OctaveEcoute, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.ValCtrl, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Lettre_G As Label
    Friend WithEvents Lettre_F As Label
    Friend WithEvents Lettre_D As Label
    Friend WithEvents Lettre_S As Label
    Friend WithEvents Lettre_H As Label
    Friend WithEvents Lettre_L As Label
    Friend WithEvents Lettre_K As Label
    Friend WithEvents Lettre_J As Label
    Friend WithEvents Lettre_I As Label
    Friend WithEvents Lettre_U As Label
    Friend WithEvents Lettre_Y As Label
    Friend WithEvents Lettre_R As Label
    Friend WithEvents Lettre_E As Label
    Friend WithEvents Panel2 As Panel
    Friend WithEvents OctaveEcoute As NumericUpDown
    Friend WithEvents Label18 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents VéloEcoute As NumericUpDown
    Friend WithEvents Label2 As Label
    Friend WithEvents ListPRG As Windows.Forms.ComboBox
    Friend WithEvents CanalEcoute As NumericUpDown
    Friend WithEvents Label1 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents VolumeEcoute As TrackBar
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Panel3 As Panel
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents ValCtrl As TrackBar
End Class
