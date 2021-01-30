using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;

namespace Jess.SmartExcel
{
    /// <summary>
    /// Excel读写类
    /// </summary>
    public class SmartExcel
    {
        //        'the memory copy API is used in the MKI$ function which converts an integer
        //        'value to a 2-byte string value to write to the file. (used by the Horizontal
        //        'Page Break function).
        //        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As String, ByRef lpvSource As Short, ByVal cbCopy As Integer)
        //
        [DllImport("kernel32.dll")]
        private static extern void RtlMoveMemory(ref string lpvDest, ref short lpvSource, int cbCopy);

        private FileStream fs;
        private BEG_FILE_RECORD m_udtBEG_FILE_MARKER;
        private END_FILE_RECORD m_udtEND_FILE_MARKER;
        private HPAGE_BREAK_RECORD m_udtHORIZ_PAGE_BREAK;


        //create an array that will hold the rows where a horizontal page break will be inserted just before.
        private int[] m_shtHorizPageBreakRows;
        private int m_shtNumHorizPageBreaks = 1;

        private void FilePut(byte[] buf)
        {
            fs.Write(buf, 0, buf.Length);
        }

        private void FilePut(System.ValueType vt)
        {
            Type t = vt.GetType();
            int size = 0;

            //            foreach(FieldInfo fi in t.GetFields())
            //            {
            //                size += Marshal.SizeOf(fi.FieldType);
            //                
            //                System.Diagnostics.Trace.WriteLine(fi.Name);
            //            }

            size = Marshal.SizeOf(vt);
            IntPtr p = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(vt, p, true);
            byte[] buf = new byte[size];
            Marshal.Copy(p, buf, 0, size);

            fs.Write(buf, 0, buf.Length);
            Marshal.FreeHGlobal(p);
        }

        private void FilePut(System.ValueType vt, int len)
        {
            int size = 0;
            size = len;
            IntPtr p = Marshal.AllocHGlobal(size);
            Marshal.StructureToPtr(vt, p, true);
            byte[] buf = new byte[size];
            Marshal.Copy(p, buf, 0, size);

            fs.Write(buf, 0, buf.Length);
            Marshal.FreeHGlobal(p);
        }

        public bool PrintGridLines
        {
            set
            {
                try
                {
                    PRINT_GRIDLINES_RECORD GRIDLINES_RECORD;

                    GRIDLINES_RECORD.opcode = 43;
                    GRIDLINES_RECORD.length = 2;
                    if (true == value)
                    {
                        GRIDLINES_RECORD.PrintFlag = 1;
                    }
                    else
                    {
                        GRIDLINES_RECORD.PrintFlag = 0;
                    }
                    FilePut(GRIDLINES_RECORD);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        public bool ProtectSpreadsheet
        {
            set
            {
                try
                {
                    PROTECT_SPREADSHEET_RECORD PROTECT_RECORD;


                    PROTECT_RECORD.opcode = 18;
                    PROTECT_RECORD.length = 2;
                    if (true == value)
                    {
                        PROTECT_RECORD.Protect = 1;
                    }
                    else
                    {
                        PROTECT_RECORD.Protect = 0;
                    }

                    if (null == fs) throw new SmartExcelException();
                    FilePut(PROTECT_RECORD);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        //
        public void CreateFile(string strFileName)
        {
            try
            {
                if (File.Exists(strFileName))
                {
                    File.SetAttributes(strFileName, FileAttributes.Normal);
                    File.Delete(strFileName);
                }

                fs = new FileStream(strFileName, FileMode.CreateNew);
                FilePut(m_udtBEG_FILE_MARKER);

                WriteDefaultFormats();
                //                    'create the Horizontal Page Break array
                m_shtHorizPageBreakRows = new int[1] { 0 };
                m_shtNumHorizPageBreaks = 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        //
        public void CloseFile()
        {
            short x;
            try
            {
                if (null != fs)
                {
                    //'write the horizontal page breaks if necessary
                    int lLoop1, lLoop2, lTemp;

                    if (m_shtNumHorizPageBreaks > 0)
                    {
                        //                                               'the Horizontal Page Break array must be in sorted order.
                        //                'Use a simple Bubble sort because the size of this array would
                        //                'be pretty small most of the time. A QuickSort would probably
                        //                'be overkill.
                        for (lLoop1 = m_shtHorizPageBreakRows.GetUpperBound(0); lLoop1 >= m_shtHorizPageBreakRows.GetLowerBound(0); lLoop1--)
                        {
                            for (lLoop2 = m_shtHorizPageBreakRows.GetLowerBound(0) + 1; lLoop2 <= lLoop1; lLoop2++)
                            {
                                if (m_shtHorizPageBreakRows[lLoop2 - 1] > m_shtHorizPageBreakRows[lLoop2])
                                {
                                    lTemp = m_shtHorizPageBreakRows[lLoop2 - 1];
                                    m_shtHorizPageBreakRows[lLoop2 - 1] = m_shtHorizPageBreakRows[lLoop2];
                                    m_shtHorizPageBreakRows[lLoop2] = (short)lTemp;
                                }
                            }
                        }
                        //'write the Horizontal Page Break Record
                        m_udtHORIZ_PAGE_BREAK.opcode = 27;
                        m_udtHORIZ_PAGE_BREAK.length = (short)(2 + (m_shtNumHorizPageBreaks * 2));
                        m_udtHORIZ_PAGE_BREAK.NumPageBreaks = (short)m_shtNumHorizPageBreaks;

                        FilePut(m_udtHORIZ_PAGE_BREAK);

                        //                                             'now write the actual page break values
                        //                'the MKI$ function is standard in other versions of BASIC but
                        //                'VisualBasic does not have it. A KnowledgeBase article explains
                        //                'how to recreate it (albeit using 16-bit API, I switched it
                        //                'to 32-bit).
                        for (x = 1; x <= m_shtHorizPageBreakRows.GetUpperBound(0); x++)
                        {
                            FilePut(System.Text.Encoding.Default.GetBytes(MKI((short)(m_shtHorizPageBreakRows[x]))));
                        }
                    }
                    FilePut(m_udtEND_FILE_MARKER);
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void Init()
        {
            //                                                                         'Set up default values for records
            //        'These should be the values that are the same for every record of these types

            // beginning of file
            m_udtBEG_FILE_MARKER.opcode = 9;
            m_udtBEG_FILE_MARKER.length = 4;
            m_udtBEG_FILE_MARKER.version = 2;
            m_udtBEG_FILE_MARKER.ftype = 10;

            // end of file marker
            m_udtEND_FILE_MARKER.opcode = 10;
        }

        public SmartExcel()
        {
            Init();
        }


        public void InsertHorizPageBreak(int lrow)
        {
            int row;
            try
            {
                //    'the row and column values are written to the excel file as
                //    'unsigned integers. Therefore, must convert the longs to integer.
                if (lrow > 32767 || lrow < 0) row = 0;
                else row = lrow - 1;
                m_shtNumHorizPageBreaks = m_shtNumHorizPageBreaks + 1;
                m_shtHorizPageBreakRows[m_shtNumHorizPageBreaks] = row;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void WriteValue(ValueTypes ValueType, CellFont CellFontUsed, CellAlignment Alignment, CellHiddenLocked HiddenLocked, int lrow, int lcol, object Value)
        {
            WriteValue(ValueType, CellFontUsed, Alignment, HiddenLocked, lrow, lcol, Value, 0);
        }

        public void WriteValue(ValueTypes ValueType, CellFont CellFontUsed, CellAlignment Alignment, CellHiddenLocked HiddenLocked, int lrow, int lcol, object Value, int CellFormat)
        {
            int l;
            string st;
            short col, row;
            try
            {
                //                            'the row and column values are written to the excel file as
                //                'unsigned integers. Therefore, must convert the longs to integer.
                tInteger INTEGER_RECORD;
                tNumber NUMBER_RECORD;
                byte b;
                tText TEXT_RECORD;
                if (lrow > 32767 || lrow < 0) row = 0;
                else row = (short)(lrow - 1);
                if (lcol > 32767 || lcol < 0) col = 0; else col = (short)(lcol - 1);
                switch (ValueType)
                {
                    case ValueTypes.Integer:
                        INTEGER_RECORD.opcode = 2;
                        INTEGER_RECORD.length = 9;
                        INTEGER_RECORD.row = row;
                        INTEGER_RECORD.col = col;
                        INTEGER_RECORD.rgbAttr1 = (byte)(HiddenLocked);
                        INTEGER_RECORD.rgbAttr2 = (byte)(CellFontUsed + CellFormat);
                        INTEGER_RECORD.rgbAttr3 = (byte)(Alignment);
                        INTEGER_RECORD.intValue = (short)(Value);
                        FilePut(INTEGER_RECORD);
                        break;
                    case ValueTypes.Number:
                        NUMBER_RECORD.opcode = 3;
                        NUMBER_RECORD.length = 15;
                        NUMBER_RECORD.row = row;
                        NUMBER_RECORD.col = col;
                        NUMBER_RECORD.rgbAttr1 = (byte)(HiddenLocked);
                        NUMBER_RECORD.rgbAttr2 = (byte)(CellFontUsed + CellFormat);
                        NUMBER_RECORD.rgbAttr3 = (byte)(Alignment);
                        NUMBER_RECORD.NumberValue = (double)(Value);
                        FilePut(NUMBER_RECORD);
                        break;
                    case ValueTypes.Text:
                        st = Convert.ToString(Value);
                        l = GetLength(st);// 'LenB(StrConv(st, vbFromUnicode)) 'Len(st$)

                        TEXT_RECORD.opcode = 4;
                        TEXT_RECORD.length = 10;
                        //'Length of the text portion of the record
                        TEXT_RECORD.TextLength = (byte)l;
                        //                          'Total length of the record
                        TEXT_RECORD.length = (byte)(8 + l);
                        TEXT_RECORD.row = row;
                        TEXT_RECORD.col = col;
                        TEXT_RECORD.rgbAttr1 = (byte)(HiddenLocked);
                        TEXT_RECORD.rgbAttr2 = (byte)(CellFontUsed + CellFormat);
                        TEXT_RECORD.rgbAttr3 = (byte)(Alignment);
                        //                        'Put record header
                        FilePut(TEXT_RECORD);
                        //                'Then the actual string data
                        FilePut(System.Text.Encoding.Default.GetBytes(st));
                        break;
                    default: break;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetMargin(MarginTypes Margin, double MarginValue)
        {
            try
            {
                //                      'write the spreadsheet's layout information (in inches)
                MARGIN_RECORD_LAYOUT MarginRecord;
                MarginRecord.opcode = (short)Margin;
                MarginRecord.length = 8;
                MarginRecord.MarginValue = MarginValue;// 'in inches
                FilePut(MarginRecord);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetColumnWidth(int FirstColumn, int LastColumn, short WidthValue)
        {
            try
            {
                COLWIDTH_RECORD COLWIDTH;

                COLWIDTH.opcode = 36;
                COLWIDTH.length = 4;
                COLWIDTH.col1 = (byte)(FirstColumn - 1);
                COLWIDTH.col2 = (byte)(LastColumn - 1);
                COLWIDTH.ColumnWidth = (short)(WidthValue * 256);// 'values are specified as 1/256 of a character
                FilePut(COLWIDTH);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetFont(string FontName, short FontHeight, FontFormatting FontFormat)
        {
            int l;
            try
            {
                //     'you can set up to 4 fonts in the spreadsheet file. When writing a value such
                //    'as a Text or Number you can specify one of the 4 fonts (numbered 0 to 3)
                FONT_RECORD FONTNAME_RECORD;
                l = GetLength(FontName);// 'LenB(StrConv(FontName, vbFromUnicode)) 'Len(FontName)
                FONTNAME_RECORD.opcode = 49;
                FONTNAME_RECORD.length = (short)(5 + l);
                FONTNAME_RECORD.FontHeight = (short)(FontHeight * 20);
                FONTNAME_RECORD.FontAttributes1 = (byte)FontFormat;// 'bold/underline etc
                FONTNAME_RECORD.FontAttributes2 = (byte)0;// 'reserved-always zero!!
                FONTNAME_RECORD.FontNameLength = (byte)l;//'CByte(Len(FontName))
                FilePut(FONTNAME_RECORD);
                //                        'Then the actual font name data
                FilePut(System.Text.Encoding.Default.GetBytes(FontName));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetHeader(string HeaderText)
        {
            int l;
            try
            {
                HEADER_FOOTER_RECORD HEADER_RECORD;
                l = GetLength(HeaderText);//   'LenB(StrConv(HeaderText, vbFromUnicode)) 'Len(HeaderText)
                HEADER_RECORD.opcode = 20;
                HEADER_RECORD.length = (short)(1 + l);
                HEADER_RECORD.TextLength = (byte)l;// 'CByte(Len(HeaderText))
                FilePut(HEADER_RECORD);
                //                        'Then the actual Header text
                FilePut(System.Text.Encoding.Default.GetBytes(HeaderText));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetFooter(string FooterText)
        {
            int l;
            try
            {
                HEADER_FOOTER_RECORD FOOTER_RECORD;
                l = GetLength(FooterText);// 'LenB(StrConv(FooterText, vbFromUnicode)) 'Len(FooterText)
                FOOTER_RECORD.opcode = 21;
                FOOTER_RECORD.length = (short)(1 + l);
                FOOTER_RECORD.TextLength = (byte)l;
                FilePut(FOOTER_RECORD);
                //                    'Then the actual Header text
                FilePut(System.Text.Encoding.Default.GetBytes(FooterText));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void SetFilePassword(string PasswordText)
        {
            int l;
            try
            {
                PASSWORD_RECORD FILE_PASSWORD_RECORD;
                l = GetLength(PasswordText);// 'LenB(StrConv(PasswordText, vbFromUnicode)) 'Len(PasswordText)
                FILE_PASSWORD_RECORD.opcode = 47;
                FILE_PASSWORD_RECORD.length = (short)l;
                FilePut(FILE_PASSWORD_RECORD);
                //          'Then the actual Password text
                FilePut(System.Text.Encoding.Default.GetBytes(PasswordText));
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void WriteDefaultFormats()
        {
            FORMAT_COUNT_RECORD cFORMAT_COUNT_RECORD;
            FORMAT_RECORD cFORMAT_RECORD;
            int lIndex;
            int l;
            string q = "\"";
            string[] aFormat = new string[]{
                                               "General",
                                               "0",
                                               "0.00",
                                               "#,##0",
                                               "#,##0.00",
                                               "#,##0\\ "+q+"$"+q+";\\-#,##0\\ "+q+"$"+q,
                                               "#,##0\\ "+q+"$"+q+";[Red]\\-#,##0\\ "+q+"$"+q,
                                               "#,##0.00\\ "+q+"$"+q+";\\-#,##0.00\\ "+q+"$"+q,
                                               "#,##0.00\\ "+q+"$"+q+";[Red]\\-#,##0.00\\ "+q+"$"+q,
                                               "0%",
                                               "0.00%",
                                               "0.00E+00",
                                               "dd/mm/yy",
                                               "dd/\\ mmm\\ yy",
                                               "dd/\\ mmm",
                                               "mmm\\ yy",
                                               "h:mm\\ AM/PM",
                                               "h:mm:ss\\ AM/PM",
                                               "hh:mm",
                                               "hh:mm:ss",
                                               "dd/mm/yy\\ hh:mm",
                                               "##0.0E+0",
                                               "mm:ss",
                                               "@"};


            cFORMAT_COUNT_RECORD.opcode = 0x1f;
            cFORMAT_COUNT_RECORD.length = 0x02;
            cFORMAT_COUNT_RECORD.Count = (short)(aFormat.GetUpperBound(0));
            FilePut(cFORMAT_COUNT_RECORD);

            byte b;
            int a;
            for (lIndex = aFormat.GetLowerBound(0); lIndex <= aFormat.GetUpperBound(0); lIndex++)
            {
                l = aFormat[lIndex].Length;
                cFORMAT_RECORD.opcode = 0x1e;
                cFORMAT_RECORD.length = (short)(l + 1);
                cFORMAT_RECORD.FormatLength = (byte)(l);
                FilePut(cFORMAT_RECORD);
                //                Then the actual format
                // 从1开始还是从0开始？！
                for (a = 0; a < l; a++)
                {
                    b = (byte)(aFormat[lIndex].Substring(a, 1).ToCharArray(0, 1)[0]);
                    FilePut(new byte[] { b });
                }
            }
        }

        private string MKI(short x)
        {
            string temp;
            //'used for writing integer array values to the disk file
            temp = "  ";
            RtlMoveMemory(ref temp, ref x, 2);
            return temp;
        }

        private int GetLength(string strText)
        {
            return System.Text.Encoding.Default.GetBytes(strText).Length;
        }

        public void SetDefaultRowHeight(int HeightValue)
        {
            try
            {
                //              'Height is defined in units of 1/20th of a point. Therefore, a 10-point font
                //                'would be 200 (i.e. 200/20 = 10). This function takes a HeightValue such as
                //                '14 point and converts it the correct size before writing it to the file.

                DEF_ROWHEIGHT_RECORD DEFHEIGHT;
                DEFHEIGHT.opcode = 37;
                DEFHEIGHT.length = 2;
                DEFHEIGHT.RowHeight = HeightValue * 20;//  'convert points to 1/20ths of point
                FilePut(DEFHEIGHT);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SetRowHeight(int Row, short HeightValue)
        {
            int o_intRow;
            //               'the row and column values are written to the excel file as
            //                'unsigned integers. Therefore, must convert the longs to integer.
            if (Row > 32767) throw new Exception("行号不能大于32767！");
            try
            {
                o_intRow = Row;
                //                if(Row > 32767){
                //                                   o_intRow = CInt(Row - 65536)
                //                                   Else
                //                                       o_intRow = CInt(Row) - 1    'rows/cols in Excel binary file are zero based
                //                                                                   }

                //                    'Height is defined in units of 1/20th of a point. Therefore, a 10-point font
                //                'would be 200 (i.e. 200/20 = 10). This function takes a HeightValue such as
                //                '14 point and converts it the correct size before writing it to the file.

                ROW_HEIGHT_RECORD ROWHEIGHTREC;
                ROWHEIGHTREC.opcode = 8;
                ROWHEIGHTREC.length = 16;
                ROWHEIGHTREC.RowNumber = o_intRow;
                ROWHEIGHTREC.FirstColumn = 0;
                ROWHEIGHTREC.LastColumn = 256;
                ROWHEIGHTREC.RowHeight = HeightValue * 20;// 'convert points to 1/20ths of point
                ROWHEIGHTREC.internals = 0;
                ROWHEIGHTREC.DefaultAttributes = 0;
                ROWHEIGHTREC.FileOffset = 0;
                ROWHEIGHTREC.rgbAttr1 = 0;
                ROWHEIGHTREC.rgbAttr2 = 0;
                ROWHEIGHTREC.rgbAttr3 = 0;
                FilePut(ROWHEIGHTREC);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}