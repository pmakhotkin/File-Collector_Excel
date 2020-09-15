
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
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

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxPath = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Algerian", 13.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(451, 26)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Введите путь к папке с исходными файлами :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'TextBoxPath
        '
        Me.TextBoxPath.Location = New System.Drawing.Point(19, 73)
        Me.TextBoxPath.Name = "TextBoxPath"
        Me.TextBoxPath.Size = New System.Drawing.Size(443, 22)
        Me.TextBoxPath.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(21, 177)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(443, 22)
        Me.TextBox1.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Algerian", 13.8!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(25, 141)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(425, 26)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Введите путь для сохранения результата : "
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(103, 252)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(270, 38)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Собрать файлы"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.CausesValidation = False
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(22, 304)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(440, 61)
        Me.Label3.TabIndex = 5
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label3.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(514, 372)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxPath)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "File Collector"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents TextBoxPath As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Button1 As Button



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBoxPath.Text = "" Or TextBox1.Text = "" Then
            Label3.Text = "Введите пути для продолжения"
            Label3.Visible = True
        End If
        If TextBoxPath.Text <> "" And TextBox1.Text <> "" Then
            Label3.Visible = False
            If IO.Directory.Exists(TextBoxPath.Text) = False Then
                Label3.Text = "папка не найдена, проверьте путь для исходных файлов"
                Label3.Visible = True
            Else
                If IO.Directory.Exists(TextBox1.Text) = False Then
                    Label3.Text = "папка не найдена, проверьте путь для coхранения результата"
                    Label3.Visible = True
                End If
            End If

        End If

        If IO.Directory.Exists(TextBoxPath.Text) = True And IO.Directory.Exists(TextBox1.Text) = True Then
            Dim checkfile As Integer = IO.Directory.GetFiles(TextBoxPath.Text, "*.x*").Length

            If checkfile = 0 Then
                Label3.Text = "В заданной папке нет файлов Excel"
                Label3.Visible = True
            Else
                'создаем шаблон итогового файла

                Dim oExcel As Object
                Dim oBook As Object
                Dim oSheet As Object
                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.Workbooks.Add
                oSheet = oBook.Worksheets(1)

                With oSheet
                    .Name = "перевод"
                    .Range("A7") = "№ Инвойса"
                    .Range("B7") = "Артикул"
                    .Range("C7") = "№ ТД"
                    .Range("D7") = "ТНВЭД"
                    .Range("E7") = "№ товара в ТД"
                    .Range("F7") = "Страна (букв. код)"
                    .Range("G7") = "Наименование"
                    .Range("A7:G7").Font.FontStyle = "Arial"
                    .Range("A7:G7").Font.Size = 12
                    .Range("A7:G7").EntireColumn.AutoFit
                End With

                oExcel.Visible = True

                ' перебираем файлы ексель в папке
                For Each file In IO.Directory.GetFiles(TextBoxPath.Text, "*.x*")


                Next
            End If
        End If


    End Sub

    Friend WithEvents Label3 As Label
End Class
