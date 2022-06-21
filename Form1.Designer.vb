<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(424, 32)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 28)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Parcourir"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(15, 34)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(400, 22)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "D:\excel_sk.xlsx"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TextBox3)
        Me.GroupBox1.Controls.Add(Me.TextBox2)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 143)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(216, 89)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Création Simple"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(103, 60)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 17)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Ligne de fin"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(103, 28)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(108, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Ligne de départ"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(8, 57)
        Me.TextBox3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBox3.Size = New System.Drawing.Size(85, 22)
        Me.TextBox3.TabIndex = 1
        Me.TextBox3.Text = "20"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(8, 25)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(85, 22)
        Me.TextBox2.TabIndex = 0
        Me.TextBox2.Text = "2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 11)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(171, 17)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Chemin vers tableur Excel"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(424, 91)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 28)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "Parcourir"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(15, 94)
        Me.TextBox4.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(400, 22)
        Me.TextBox4.TabIndex = 1
        Me.TextBox4.Text = "C:\Users\pinty\Desktop\Fichiers CAO Temp"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(16, 70)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(134, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Répertoire de sortie"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CheckBox3)
        Me.GroupBox2.Controls.Add(Me.CheckBox2)
        Me.GroupBox2.Controls.Add(Me.TextBox5)
        Me.GroupBox2.Location = New System.Drawing.Point(240, 143)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(284, 89)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Création Mulitple"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Enabled = False
        Me.CheckBox2.Location = New System.Drawing.Point(9, 28)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(145, 21)
        Me.CheckBox2.TabIndex = 13
        Me.CheckBox2.Text = "Créer assemblage"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'TextBox5
        '
        Me.TextBox5.Enabled = False
        Me.TextBox5.Location = New System.Drawing.Point(8, 57)
        Me.TextBox5.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(265, 22)
        Me.TextBox5.TabIndex = 12
        Me.TextBox5.Text = "2;2;3;20"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(424, 244)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "Générer"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(336, 244)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 17)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Patientez..."
        Me.Label4.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(15, 242)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(117, 21)
        Me.CheckBox1.TabIndex = 12
        Me.CheckBox1.Text = "Mode multiple"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Enabled = False
        Me.CheckBox3.Location = New System.Drawing.Point(162, 28)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(100, 21)
        Me.CheckBox3.TabIndex = 13
        Me.CheckBox3.Text = "Auto Parse"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(536, 282)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.RightToLeftLayout = True
        Me.Text = "SKBuilder"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TextBox3 As Windows.Forms.TextBox
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Button3 As Windows.Forms.Button
    Friend WithEvents TextBox4 As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents CheckBox2 As Windows.Forms.CheckBox
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents TextBox5 As Windows.Forms.TextBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents BindingSource1 As Windows.Forms.BindingSource
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As Windows.Forms.CheckBox
End Class
