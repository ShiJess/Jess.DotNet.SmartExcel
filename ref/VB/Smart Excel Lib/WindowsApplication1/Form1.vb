Imports COM.Excel

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows"

    Public Sub New()
        MyBase.New()

        InitializeComponent()

    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private components As System.ComponentModel.IContainer

    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(107, 82)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(181, 62)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Simple"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(107, 195)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(181, 62)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Complex"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 18)
        Me.ClientSize = New System.Drawing.Size(405, 354)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a As New cExcelFile
        With a
            .CreateFile(".\abc.xls")
            .PrintGridLines = False

            .SetMargin(cExcelFile.MarginTypes.xlsTopMargin, 1.5)   'set to 1.5 inches
            .SetMargin(cExcelFile.MarginTypes.xlsLeftMargin, 1.5)
            .SetMargin(cExcelFile.MarginTypes.xlsRightMargin, 1.5)
            .SetMargin(cExcelFile.MarginTypes.xlsBottomMargin, 1.5)

            .SetFont("Arial", "12", cExcelFile.FontFormatting.xlsItalic)

            .SetColumnWidth(1, 12, 18)

            .SetHeader("This is the header")
            .SetFooter("This ia the footer")

            .WriteValue(cExcelFile.ValueTypes.xlsText, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsCentreAlign, cExcelFile.CellHiddenLocked.xlsNormal, 1, 1, "БъЬт")

            .CloseFile()
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim myExcelFile As New cExcelFile
        Dim strFileName As String

        With myExcelFile
            'Create the new spreadsheet
            strFileName = ".\vbtest.xls"  'create spreadsheet in the current directory
            .CreateFile(strFileName)

            'set a Password for the file. If set, the rest of the spreadsheet will
            'be encrypted. If a password is used it must immediately follow the
            'CreateFile method.
            'This is different then protecting the spreadsheet (see below).
            'NOTE: For some reason this function does not work. Excel will
            'recognize that the file is password protected, but entering the password
            'will not work. Also, the file is not encrypted. Therefore, do not use
            'this function until I can figure out why it doesn't work. There is not
            'much documentation on this function available.
            '.SetFilePassword "PAUL"

            'specify whether to print the gridlines or not
            'this should come before the setting of fonts and margins
            .PrintGridLines = False

            'it is a good idea to set margins, fonts and column widths
            'prior to writing any text/numerics to the spreadsheet. These
            'should come before setting the fonts.

            .SetMargin(cExcelFile.MarginTypes.xlsTopMargin, 1.5)   'set to 1.5 inches
            .SetMargin(cExcelFile.MarginTypes.xlsLeftMargin, 1.5)
            .SetMargin(cExcelFile.MarginTypes.xlsRightMargin, 1.5)
            .SetMargin(cExcelFile.MarginTypes.xlsBottomMargin, 1.5)

            'to insert a Horizontal Page Break you need to specify the row just
            'after where you want the page break to occur. You can insert as many
            'page breaks as you wish (in any order).
            '.InsertHorizPageBreak(10)
            '.InsertHorizPageBreak(20)

            'set a default row height for the entire spreadsheet (1/20th of a point)
            '.SetDefaultRowHeight(14)

            'Up to 4 fonts can be specified for the spreadsheet. This is a
            'limitation of the Excel 2.1 format. For each value written to the
            'spreadsheet you can specify which font to use.

            .SetFont("Arial", 10, cExcelFile.FontFormatting.xlsNoFormat)             'font0
            .SetFont("Arial", 10, cExcelFile.FontFormatting.xlsBold)                 'font1
            .SetFont("Arial", 10, cExcelFile.FontFormatting.xlsBold + cExcelFile.FontFormatting.xlsUnderline)  'font2
            .SetFont("Times New Roman", 12, cExcelFile.FontFormatting.xlsItalic)             'font3

            'Column widths are specified in Excel as 1/256th of a character.
            .SetColumnWidth(1, 5, 18)

            'Set special row heights for row 1 and 2
            '.SetRowHeight(1, 30)
            '.SetRowHeight(2, 30)

            'set any header or footer that you want to print on
            'every page. This text will be centered at the top and/or
            'bottom of each page. The font will always be the font that
            'is specified as font0, therefore you should only set the
            'header/footer after specifying the fonts through SetFont.
            .SetHeader("This is the header!")
            .SetFooter("This is the footer!")

            'write some data to the spreadsheet
            .WriteValue(cExcelFile.ValueTypes.xlsInteger, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsLeftAlign, cExcelFile.CellHiddenLocked.xlsNormal, 6, 1, 20)

            'write a cell with a shaded number with a bottom border
            .WriteValue(cExcelFile.ValueTypes.xlsNumber, cExcelFile.CellFont.xlsFont1, cExcelFile.CellAlignment.xlsRightAlign + cExcelFile.CellAlignment.xlsBottomBorder + cExcelFile.CellAlignment.xlsShaded, cExcelFile.CellHiddenLocked.xlsNormal, 7, 1, 123.456)

            'write a normal left aligned string using font2 (bold & underline)
            .WriteValue(cExcelFile.ValueTypes.xlsText, cExcelFile.CellFont.xlsFont2, cExcelFile.CellAlignment.xlsLeftAlign, cExcelFile.CellHiddenLocked.xlsNormal, 8, 1, "demo string !")

            'write a locked cell. The cell will not be able to be overwritten, BUT you
            'must set the sheet PROTECTION to on before it will take effect!!!
            .WriteValue(cExcelFile.ValueTypes.xlsText, cExcelFile.CellFont.xlsFont3, cExcelFile.CellAlignment.xlsLeftAlign, cExcelFile.CellHiddenLocked.xlsLocked, 9, 1, "this field is locked!")

            'fill the cell with "F"'s
            .WriteValue(cExcelFile.ValueTypes.xlsText, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsFillCell, cExcelFile.CellHiddenLocked.xlsNormal, 10, 1, "F")

            'write a hidden cell to the spreadsheet. This only works for cells
            'that contain formulae. Text, Number, Integer value text can not be hidden
            'using this feature. It is included here for the sake of completeness.
            .WriteValue(cExcelFile.ValueTypes.xlsText, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsCentreAlign, cExcelFile.CellHiddenLocked.xlsHidden, 11, 1, "If this were a formula it would be hidden!")

            'write some dates to the file. NOTE: you need to write dates as xlsNumber
            Dim d As Date
            d = CDate("01/15/2001")
            .WriteValue(cExcelFile.ValueTypes.xlsNumber, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsCentreAlign, cExcelFile.CellHiddenLocked.xlsNormal, 15, 1, d, 12)

            d = CDate("12/31/1999")
            .WriteValue(cExcelFile.ValueTypes.xlsNumber, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsCentreAlign, cExcelFile.CellHiddenLocked.xlsNormal, 16, 1, d, 12)

            d = CDate("04/01/2002")
            .WriteValue(cExcelFile.ValueTypes.xlsNumber, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsCentreAlign, cExcelFile.CellHiddenLocked.xlsNormal, 17, 1, d, 12)

            d = CDate("10/21/1998")
            .WriteValue(cExcelFile.ValueTypes.xlsNumber, cExcelFile.CellFont.xlsFont0, cExcelFile.CellAlignment.xlsCentreAlign, cExcelFile.CellHiddenLocked.xlsNormal, 18, 1, d, 12)


            'PROTECT the spreadsheet so any cells specified as LOCKED will not be
            'overwritten. Also, all cells with HIDDEN set will hide their formulae.
            'PROTECT does not use a password.
            .ProtectSpreadsheet = True

            'Finally, close the spreadsheet
            .CloseFile()

            MsgBox("Excel BIFF Spreadsheet created." & vbCrLf & "Filename: " & strFileName, vbInformation + vbOKOnly, "Excel Class")

        End With

    End Sub
End Class
