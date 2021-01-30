Imports System.Text
Imports System.IO

Public Class cExcelFile
    'Class file for writing Microsoft Excel BIFF 2.1 files.

    'This class is intended for users who do not want to use the huge
    'Jet or ADO providers if they only want to export their data to
    'an Excel compatible file.

    'Newer versions of Excel use the OLE Structure Storage methods
    'which are quite complicated.

    'Paul Squires, November 10, 2001
    'rambo2000@canada.com

    'Added default-cellformats: Dieter Hauk January 8, 2001 dieter.hauk@epost.de
    'Added default row height: Matthew Brewster November 9, 2001

    'the memory copy API is used in the MKI$ function which converts an integer
    'value to a 2-byte string value to write to the file. (used by the Horizontal
    'Page Break function).

    'enum to handle the various types of values that can be written
    'to the excel file.
    Public Enum ValueTypes
        xlsInteger = 0
        xlsNumber = 1
        xlsText = 2
    End Enum

    'enum to hold cell alignment
    Public Enum CellAlignment
        xlsGeneralAlign = 0
        xlsLeftAlign = 1
        xlsCentreAlign = 2
        xlsRightAlign = 3
        xlsFillCell = 4
        xlsLeftBorder = 8
        xlsRightBorder = 16
        xlsTopBorder = 32
        xlsBottomBorder = 64
        xlsShaded = 128
    End Enum

    'enum to handle selecting the font for the cell
    Public Enum CellFont
        'used by rgbAttr2
        'bits 0-5 handle the *picture* formatting, not bold/underline etc...
        'bits 6-7 handle the font number
        xlsFont0 = 0
        xlsFont1 = 64
        xlsFont2 = 128
        xlsFont3 = 192
    End Enum

    Public Enum CellHiddenLocked
        'used by rgbAttr1
        'bits 0-5 must be zero
        'bit 6 locked/unlocked
        'bit 7 hidden/not hidden
        xlsNormal = 0
        xlsLocked = 64
        xlsHidden = 128
    End Enum

    'set up variables to hold the spreadsheet's layout
    Public Enum MarginTypes
        xlsLeftMargin = 38
        xlsRightMargin = 39
        xlsTopMargin = 40
        xlsBottomMargin = 41
    End Enum

    Public Enum FontFormatting
        'add these enums together. For example: xlsBold + xlsUnderline
        xlsNoFormat = 0
        xlsBold = 1
        xlsItalic = 2
        xlsUnderline = 4
        xlsStrikeout = 8
    End Enum

    Private Structure FONT_RECORD
        Dim opcode As Short '49
        Dim length As Short '5+len(fontname)
        Dim FontHeight As Short
        'bit0 bold, bit1 italic, bit2 underline, bit3 strikeout, bit4-7 reserved
        Dim FontAttributes1 As Byte
        Dim FontAttributes2 As Byte 'reserved - always 0
        Dim FontNameLength As Byte
    End Structure

    Private Structure PASSWORD_RECORD
        Dim opcode As Short '47
        Dim length As Short 'len(password)
    End Structure

    Private Structure HEADER_FOOTER_RECORD
        Dim opcode As Short '20 Header, 21 Footer
        Dim length As Short '1+len(text)
        Dim TextLength As Byte
    End Structure

    Private Structure PROTECT_SPREADSHEET_RECORD
        Dim opcode As Short '18
        Dim length As Short '2
        Dim Protect As Short
    End Structure

    Private Structure FORMAT_COUNT_RECORD
        Dim opcode As Short '1f
        Dim length As Short '2
        Dim Count As Short
    End Structure

    Private Structure FORMAT_RECORD
        Dim opcode As Short '1e
        Dim length As Short '1+len(format)
        Dim FormatLenght As Byte 'len(format)
    End Structure '+ followed by the Format-Picture

    Private Structure COLWIDTH_RECORD
        Dim opcode As Short '36
        Dim length As Short '4
        Dim col1 As Byte 'first column
        Dim col2 As Byte 'last column
        Dim ColumnWidth As Short 'at 1/256th of a character
    End Structure

    'Beginning Of File record
    Private Structure BEG_FILE_RECORD
        Dim opcode As Short
        Dim length As Short
        Dim version As Short
        Dim ftype As Short
    End Structure

    'End Of File record
    Private Structure END_FILE_RECORD
        Dim opcode As Short
        Dim length As Short
    End Structure

    'true/false to print gridlines
    Private Structure PRINT_GRIDLINES_RECORD
        Dim opcode As Short
        Dim length As Short
        Dim PrintFlag As Short
    End Structure

    'Integer record
    Private Structure tInteger
        Dim opcode As Short
        Dim length As Short
        Dim row As Short 'unsigned integer
        Dim col As Short
        'rgbAttr1 handles whether cell is hidden and/or locked
        Dim rgbAttr1 As Byte
        'rgbAttr2 handles the Font# and Formatting assigned to this cell
        Dim rgbAttr2 As Byte
        'rgbAttr3 handles the Cell Alignment/borders/shading
        Dim rgbAttr3 As Byte
        Dim intValue As Short 'the actual integer value
    End Structure

    'Number record
    Private Structure tNumber
        Dim opcode As Short
        Dim length As Short
        Dim row As Short
        Dim col As Short
        Dim rgbAttr1 As Byte
        Dim rgbAttr2 As Byte
        Dim rgbAttr3 As Byte
        Dim NumberValue As Double '8 Bytes
    End Structure

    'Label (Text) record
    Private Structure tText
        Dim opcode As Short
        Dim length As Short
        Dim row As Short
        Dim col As Short
        Dim rgbAttr1 As Byte
        Dim rgbAttr2 As Byte
        Dim rgbAttr3 As Byte
        Dim TextLength As Byte
    End Structure

    Private Structure MARGIN_RECORD_LAYOUT
        Dim opcode As Short
        Dim length As Short
        Dim MarginValue As Double '8 bytes
    End Structure

    Private Structure HPAGE_BREAK_RECORD
        Dim opcode As Short
        Dim length As Short
        Dim NumPageBreaks As Short
    End Structure

    Private Structure DEF_ROWHEIGHT_RECORD
        Dim opcode As Integer
        Dim length As Integer
        Dim RowHeight As Integer
    End Structure

    Private Structure ROW_HEIGHT_RECORD
        Dim opcode As Integer  '08
        Dim length As Integer  'should always be 16 bytes
        Dim RowNumber As Integer
        Dim FirstColumn As Integer
        Dim LastColumn As Integer
        Dim RowHeight As Integer  'written to file as 1/20ths of a point
        Dim internal As Integer
        Dim DefaultAttributes As Byte  'set to zero for no default attributes
        Dim FileOffset As Integer
        Dim rgbAttr1 As Byte
        Dim rgbAttr2 As Byte
        Dim rgbAttr3 As Byte
    End Structure

    'the memory copy API is used in the MKI$ function which converts an integer
    'value to a 2-byte string value to write to the file. (used by the Horizontal
    'Page Break function).
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As String, ByRef lpvSource As Short, ByVal cbCopy As Integer)

    Private m_shtFileNumber As Short
    Private m_udtBEG_FILE_MARKER As BEG_FILE_RECORD
    Private m_udtEND_FILE_MARKER As END_FILE_RECORD
    Private m_udtHORIZ_PAGE_BREAK As HPAGE_BREAK_RECORD

    'create an array that will hold the rows where a horizontal page
    'break will be inserted just before.
    Private m_shtHorizPageBreakRows() As Short
    Private m_shtNumHorizPageBreaks As Short



    Public WriteOnly Property PrintGridLines() As Boolean
        Set(ByVal Value As Boolean)
            Try
                Dim GRIDLINES_RECORD As PRINT_GRIDLINES_RECORD

                With GRIDLINES_RECORD
                    .opcode = 43
                    .length = 2
                    If Value = True Then
                        .PrintFlag = 1
                    Else
                        .PrintFlag = 0
                    End If

                End With

                FilePut(m_shtFileNumber, GRIDLINES_RECORD)
            Catch ex As Exception

            End Try
        End Set
    End Property

    Public WriteOnly Property ProtectSpreadsheet() As Boolean
        Set(ByVal Value As Boolean)
            Try
                Dim PROTECT_RECORD As PROTECT_SPREADSHEET_RECORD

                With PROTECT_RECORD
                    .opcode = 18
                    .length = 2
                    If Value = True Then
                        .Protect = 1
                    Else
                        .Protect = 0
                    End If

                End With

                FilePut(m_shtFileNumber, PROTECT_RECORD)

            Catch ex As Exception

            End Try
        End Set
    End Property

    Public Function CreateFile(ByVal strFileName As String) As Integer
        Dim OpenFile As Integer

        Try
            If File.Exists(strFileName) Then
                File.SetAttributes(strFileName, FileAttributes.Normal)
                File.Delete(strFileName)
            End If

            m_shtFileNumber = FreeFile()

            FileOpen(m_shtFileNumber, strFileName, OpenMode.Binary)

            FilePut(m_shtFileNumber, m_udtBEG_FILE_MARKER) 'must always be written first

            Call WriteDefaultFormats()

            'create the Horizontal Page Break array
            ReDim m_shtHorizPageBreakRows(0)

            m_shtNumHorizPageBreaks = 0

            OpenFile = 0 'return with no error

        Catch ex As Exception
            OpenFile = Err.Number
        End Try

    End Function

    Public Function CloseFile() As Integer
        Dim x As Short

        Try
            If m_shtFileNumber > 0 Then
                'write the horizontal page breaks if necessary
                Dim lLoop1 As Integer
                Dim lLoop2 As Integer
                Dim lTemp As Integer
                If m_shtNumHorizPageBreaks > 0 Then
                    'the Horizontal Page Break array must be in sorted order.
                    'Use a simple Bubble sort because the size of this array would
                    'be pretty small most of the time. A QuickSort would probably
                    'be overkill.
                    For lLoop1 = UBound(m_shtHorizPageBreakRows) To LBound(m_shtHorizPageBreakRows) Step -1
                        For lLoop2 = LBound(m_shtHorizPageBreakRows) + 1 To lLoop1
                            If m_shtHorizPageBreakRows(lLoop2 - 1) > m_shtHorizPageBreakRows(lLoop2) Then
                                lTemp = m_shtHorizPageBreakRows(lLoop2 - 1)
                                m_shtHorizPageBreakRows(lLoop2 - 1) = m_shtHorizPageBreakRows(lLoop2)
                                m_shtHorizPageBreakRows(lLoop2) = lTemp
                            End If
                        Next lLoop2
                    Next lLoop1

                    'write the Horizontal Page Break Record
                    With m_udtHORIZ_PAGE_BREAK
                        .opcode = 27
                        .length = 2 + (m_shtNumHorizPageBreaks * 2)
                        .NumPageBreaks = m_shtNumHorizPageBreaks
                    End With

                    FilePut(m_shtFileNumber, m_udtHORIZ_PAGE_BREAK)

                    'now write the actual page break values
                    'the MKI$ function is standard in other versions of BASIC but
                    'VisualBasic does not have it. A KnowledgeBase article explains
                    'how to recreate it (albeit using 16-bit API, I switched it
                    'to 32-bit).
                    For x = 1 To UBound(m_shtHorizPageBreakRows)
                        FilePut(m_shtFileNumber, MKI(m_shtHorizPageBreakRows(x)))
                    Next
                End If

                FilePut(m_shtFileNumber, m_udtEND_FILE_MARKER)
                FileClose(m_shtFileNumber)

                CloseFile = 0 'return with no error code
            Else
                CloseFile = -1
            End If
        Catch ex As Exception
            CloseFile = Err.Number
        End Try

    End Function

    Private Sub Init()

        'Set up default values for records
        'These should be the values that are the same for every record of these types

        With m_udtBEG_FILE_MARKER 'beginning of file
            .opcode = 9
            .length = 4
            .version = 2
            .ftype = 10
        End With

        With m_udtEND_FILE_MARKER 'end of file marker
            .opcode = 10
        End With

    End Sub

    Public Sub New()
        MyBase.New()

        Init()
    End Sub

    Public Function InsertHorizPageBreak(ByRef lrow As Integer) As Integer
        Dim row As Short

        Try
            'the row and column values are written to the excel file as
            'unsigned integers. Therefore, must convert the longs to integer.
            If lrow > 32767 Then
                row = CShort(lrow - 65536)
            Else
                row = CShort(lrow) - 1 'rows/cols in Excel binary file are zero based
            End If

            m_shtNumHorizPageBreaks = m_shtNumHorizPageBreaks + 1
            ReDim Preserve m_shtHorizPageBreakRows(m_shtNumHorizPageBreaks)

            m_shtHorizPageBreakRows(m_shtNumHorizPageBreaks) = row

        Catch ex As Exception
            InsertHorizPageBreak = Err.Number
        End Try

    End Function

    Public Function WriteValue(ByRef ValueType As ValueTypes, ByRef CellFontUsed As CellFont, ByRef Alignment As CellAlignment, ByRef HiddenLocked As CellHiddenLocked, ByRef lrow As Integer, ByRef lcol As Integer, ByRef Value As Object, Optional ByRef CellFormat As Integer = 0) As Integer
        Dim l As Short
        Dim st As String
        Dim col As Short
        Dim row As Short

        Try
            'the row and column values are written to the excel file as
            'unsigned integers. Therefore, must convert the longs to integer.

            Dim INTEGER_RECORD As tInteger
            Dim NUMBER_RECORD As tNumber
            Dim b As Byte
            Dim TEXT_RECORD As tText

            If lrow > 32767 Then
                row = CShort(lrow - 65536)
            Else
                row = CShort(lrow) - 1 'rows/cols in Excel binary file are zero based
            End If

            If lcol > 32767 Then
                col = CShort(lcol - 65536)
            Else
                col = CShort(lcol) - 1 'rows/cols in Excel binary file are zero based
            End If

            Select Case ValueType
                Case ValueTypes.xlsInteger
                    With INTEGER_RECORD
                        .opcode = 2
                        .length = 9
                        .row = row
                        .col = col
                        .rgbAttr1 = CByte(HiddenLocked)
                        .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                        .rgbAttr3 = CByte(Alignment)
                        .intValue = CShort(Value)
                    End With

                    FilePut(m_shtFileNumber, INTEGER_RECORD)

                Case ValueTypes.xlsNumber
                    With NUMBER_RECORD
                        .opcode = 3
                        .length = 15
                        .row = row
                        .col = col
                        .rgbAttr1 = CByte(HiddenLocked)
                        .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                        .rgbAttr3 = CByte(Alignment)
                        .NumberValue = CDbl(Value)
                    End With

                    FilePut(m_shtFileNumber, NUMBER_RECORD)

                Case ValueTypes.xlsText
                    st = CType(Value, String)

                    l = GetLength(st) 'LenB(StrConv(st, vbFromUnicode)) 'Len(st$)

                    With TEXT_RECORD
                        .opcode = 4
                        .length = 10
                        'Length of the text portion of the record
                        .TextLength = l

                        'Total length of the record
                        .length = 8 + l

                        .row = row
                        .col = col

                        .rgbAttr1 = CByte(HiddenLocked)
                        .rgbAttr2 = CByte(CellFontUsed + CellFormat)
                        .rgbAttr3 = CByte(Alignment)

                        'Put record header
                        FilePut(m_shtFileNumber, TEXT_RECORD)

                        'Then the actual string data
                        'For a = 1 To l%
                        '   b = Asc(Mid$(st$, a, 1))
                        '   Put #m_shtFileNumber, , b
                        'Next

                        FilePut(m_shtFileNumber, st)
                    End With

            End Select

            WriteValue = 0 'return with no error
        Catch ex As Exception
            WriteValue = Err.Number
        End Try

    End Function

    Public Function SetMargin(ByRef Margin As MarginTypes, ByRef MarginValue As Double) As Integer

        Try
            'write the spreadsheet's layout information (in inches)
            Dim MarginRecord As MARGIN_RECORD_LAYOUT

            With MarginRecord
                .opcode = Margin
                .length = 8
                .MarginValue = MarginValue 'in inches
            End With

            FilePut(m_shtFileNumber, MarginRecord)

            SetMargin = 0

        Catch ex As Exception
            SetMargin = Err.Number
        End Try

    End Function

    Public Function SetColumnWidth(ByRef FirstColumn As Byte, ByRef LastColumn As Byte, ByRef WidthValue As Short) As Integer
        Try
            Dim COLWIDTH As COLWIDTH_RECORD

            With COLWIDTH
                .opcode = 36
                .length = 4
                .col1 = FirstColumn - 1
                .col2 = LastColumn - 1
                .ColumnWidth = WidthValue * 256 'values are specified as 1/256 of a character
            End With

            FilePut(m_shtFileNumber, COLWIDTH)

            SetColumnWidth = 0
        Catch ex As Exception
            SetColumnWidth = Err.Number
        End Try
    End Function

    Public Function SetFont(ByRef FontName As String, ByRef FontHeight As Short, ByRef FontFormat As FontFormatting) As Short
        Dim l As Short

        Try
            'you can set up to 4 fonts in the spreadsheet file. When writing a value such
            'as a Text or Number you can specify one of the 4 fonts (numbered 0 to 3)

            Dim FONTNAME_RECORD As FONT_RECORD

            l = GetLength(FontName) 'LenB(StrConv(FontName, vbFromUnicode)) 'Len(FontName)

            With FONTNAME_RECORD
                .opcode = 49
                .length = 5 + l
                .FontHeight = FontHeight * 20
                .FontAttributes1 = CByte(FontFormat) 'bold/underline etc...
                .FontAttributes2 = CByte(0) 'reserved-always zero!!
                .FontNameLength = CByte(l) 'CByte(Len(FontName))
            End With

            FilePut(m_shtFileNumber, FONTNAME_RECORD)

            'Then the actual font name data
            'Dim b As Byte
            'For a = 1 To l%
            '   b = Asc(Mid$(FontName, a, 1))
            '   Put #m_shtFileNumber, , b
            'Next

            FilePut(m_shtFileNumber, FontName)

            SetFont = 0

        Catch ex As Exception
            SetFont = Err.Number
        End Try

    End Function

    Public Function SetHeader(ByRef HeaderText As String) As Integer
        Dim l As Short

        Try

            Dim HEADER_RECORD As HEADER_FOOTER_RECORD

            l = GetLength(HeaderText)   'LenB(StrConv(HeaderText, vbFromUnicode)) 'Len(HeaderText)

            With HEADER_RECORD
                .opcode = 20
                .length = 1 + l
                .TextLength = CByte(l) 'CByte(Len(HeaderText))
            End With

            FilePut(m_shtFileNumber, HEADER_RECORD)

            'Then the actual Header text
            'Dim b As Byte
            'For a = 1 To l%
            '   b = Asc(Mid$(HeaderText, a, 1))
            '   Put #m_shtFileNumber, , b
            'Next

            FilePut(m_shtFileNumber, HeaderText)

            SetHeader = 0

        Catch ex As Exception
            SetHeader = Err.Number
        End Try

    End Function

    Public Function SetFooter(ByRef FooterText As String) As Integer
        Dim l As Short

        Try
            Dim FOOTER_RECORD As HEADER_FOOTER_RECORD

            l = GetLength(FooterText) 'LenB(StrConv(FooterText, vbFromUnicode)) 'Len(FooterText)

            With FOOTER_RECORD
                .opcode = 21
                .length = 1 + l
                .TextLength = CByte(l) 'CByte(Len(FooterText))
            End With

            FilePut(m_shtFileNumber, FOOTER_RECORD)

            'Then the actual Header text
            'Dim b As Byte
            'For a = 1 To l%
            '   b = Asc(Mid$(FooterText, a, 1))
            '   Put #m_shtFileNumber, , b
            'Next

            FilePut(m_shtFileNumber, FooterText)

            SetFooter = 0

        Catch ex As Exception
            SetFooter = Err.Number
        End Try

    End Function

    Public Function SetFilePassword(ByRef PasswordText As String) As Integer
        Dim l As Short

        Try
            Dim FILE_PASSWORD_RECORD As PASSWORD_RECORD

            l = GetLength(PasswordText) 'LenB(StrConv(PasswordText, vbFromUnicode)) 'Len(PasswordText)

            With FILE_PASSWORD_RECORD
                .opcode = 47
                .length = l
            End With

            FilePut(m_shtFileNumber, FILE_PASSWORD_RECORD)

            'Then the actual Password text
            'Dim b As Byte
            'For a = 1 To l%
            '   b = Asc(Mid$(PasswordText, a, 1))
            '   Put #m_shtFileNumber, , b
            'Next

            FilePut(m_shtFileNumber, PasswordText)

            SetFilePassword = 0

        Catch ex As Exception
            SetFilePassword = Err.Number
        End Try

    End Function

    Private Function WriteDefaultFormats() As Integer

        Dim cFORMAT_COUNT_RECORD As FORMAT_COUNT_RECORD
        Dim cFORMAT_RECORD As FORMAT_RECORD
        Dim lIndex As Integer
        Dim aFormat(23) As String
        Dim l As Integer
        Dim q As String = Chr(34)

        aFormat(0) = "General"
        aFormat(1) = "0"
        aFormat(2) = "0.00"
        aFormat(3) = "#,##0"
        aFormat(4) = "#,##0.00"
        aFormat(5) = "#,##0\ " & q & "$" & q & ";\-#,##0\ " & q & "$" & q
        aFormat(6) = "#,##0\ " & q & "$" & q & ";[Red]\-#,##0\ " & q & "$" & q
        aFormat(7) = "#,##0.00\ " & q & "$" & q & ";\-#,##0.00\ " & q & "$" & q
        aFormat(8) = "#,##0.00\ " & q & "$" & q & ";[Red]\-#,##0.00\ " & q & "$" & q
        aFormat(9) = "0%"
        aFormat(10) = "0.00%"
        aFormat(11) = "0.00E+00"
        aFormat(12) = "dd/mm/yy"
        aFormat(13) = "dd/\ mmm\ yy"
        aFormat(14) = "dd/\ mmm"
        aFormat(15) = "mmm\ yy"
        aFormat(16) = "h:mm\ AM/PM"
        aFormat(17) = "h:mm:ss\ AM/PM"
        aFormat(18) = "hh:mm"
        aFormat(19) = "hh:mm:ss"
        aFormat(20) = "dd/mm/yy\ hh:mm"
        aFormat(21) = "##0.0E+0"
        aFormat(22) = "mm:ss"
        aFormat(23) = "@"

        With cFORMAT_COUNT_RECORD
            .opcode = &H1FS
            .length = &H2S
            .Count = CShort(UBound(aFormat))
        End With

        FilePut(m_shtFileNumber, cFORMAT_COUNT_RECORD)

        Dim b As Byte
        Dim a As Integer
        For lIndex = LBound(aFormat) To UBound(aFormat)
            l = Len(aFormat(lIndex))
            With cFORMAT_RECORD
                .opcode = &H1ES
                .length = CShort(l + 1)
                .FormatLenght = CShort(l)
            End With

            FilePut(m_shtFileNumber, cFORMAT_RECORD)

            'Then the actual format
            For a = 1 To l
                b = Asc(Mid(aFormat(lIndex), a, 1))
                FilePut(m_shtFileNumber, b)
            Next
        Next lIndex

    End Function

    Private Function MKI(ByRef x As Short) As String
        Dim temp As String
        'used for writing integer array values to the disk file
        temp = Space(2)
        CopyMemory(temp, x, 2)
        MKI = temp
    End Function

    Private Function GetLength(ByVal strText As String) As Integer
        Return Encoding.Default.GetBytes(strText).Length
    End Function

    Public Function SetDefaultRowHeight(ByVal HeightValue As Integer) As Integer
        Try
            'Height is defined in units of 1/20th of a point. Therefore, a 10-point font
            'would be 200 (i.e. 200/20 = 10). This function takes a HeightValue such as
            '14 point and converts it the correct size before writing it to the file.

            Dim DEFHEIGHT As DEF_ROWHEIGHT_RECORD

            With DEFHEIGHT
                .opcode = 37
                .length = 2
                .RowHeight = HeightValue * 20  'convert points to 1/20ths of point
            End With

            FilePut(m_shtFileNumber, DEFHEIGHT)

            SetDefaultRowHeight = 0

        Catch ex As Exception
            SetDefaultRowHeight = Err.Number
        End Try
    End Function

    Public Function SetRowHeight(ByVal Row As Integer, ByVal HeightValue As Short) As Integer

        Dim o_intRow As Integer

        Try
            'the row and column values are written to the excel file as
            'unsigned integers. Therefore, must convert the longs to integer.

            If Row > 32767 Then
                o_intRow = CInt(Row - 65536)
            Else
                o_intRow = CInt(Row) - 1    'rows/cols in Excel binary file are zero based
            End If

            'Height is defined in units of 1/20th of a point. Therefore, a 10-point font
            'would be 200 (i.e. 200/20 = 10). This function takes a HeightValue such as
            '14 point and converts it the correct size before writing it to the file.

            Dim ROWHEIGHTREC As ROW_HEIGHT_RECORD

            With ROWHEIGHTREC
                .opcode = 8
                .length = 16
                .RowNumber = o_intRow
                .FirstColumn = 0
                .LastColumn = 256
                .RowHeight = HeightValue * 20 'convert points to 1/20ths of point
                .internal = 0
                .DefaultAttributes = 0
                .FileOffset = 0
                .rgbAttr1 = 0
                .rgbAttr2 = 0
                .rgbAttr3 = 0
            End With

            FilePut(m_shtFileNumber, ROWHEIGHTREC)

            SetRowHeight = 0

        Catch ex As Exception
            SetRowHeight = Err.Number
        End Try
    End Function

End Class