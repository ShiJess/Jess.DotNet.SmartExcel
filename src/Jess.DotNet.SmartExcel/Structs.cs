using System;
using System.Runtime.InteropServices;

/// <summary>
/// 读写EXCEL文件所需的所有结构体的定义
/// </summary>

namespace Jess.SmartExcel
{


    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct FONT_RECORD
    {
        public short opcode;//49
        public short length;//5+len(fontname)
        public short FontHeight;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte FontAttributes1;//bit0 bold, bit1 italic, bit2 underline, bit3 strikeout, bit4-7 reserved
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte FontAttributes2;//reserved - always 0
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte FontNameLength;
    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct PASSWORD_RECORD
    {
        public short opcode;//47
        public short length;//len(password)
    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct HEADER_FOOTER_RECORD
    {
        public short opcode;//20 Header, 21 Footer
        public short length;//1+len(text)
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte TextLength;
    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct PROTECT_SPREADSHEET_RECORD
    {
        public short opcode;//18
        public short length;//2
        public short Protect;
    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct FORMAT_COUNT_RECORD
    {
        public short opcode;//0x1f
        public short length;//2
        public short Count;
    }

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct FORMAT_RECORD
    {
        public short opcode;// 0x1e
        public short length;//1+len(format)
        [MarshalAs(UnmanagedType.U1, SizeConst = 1)] public byte FormatLength;//len(format)
    }//followed by the Format-Picture

    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct COLWIDTH_RECORD
    {
        public short opcode;//36
        public short length;//4
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte col1;//first column
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte col2;//last column
        public short ColumnWidth;//at 1/256th of a character
    }

    // 'Beginning Of File record
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct BEG_FILE_RECORD
    {
        public short opcode;
        public short length;
        public short version;
        public short ftype;
    }

    // 'End Of File record
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct END_FILE_RECORD
    {
        public short opcode;
        public short length;
    }


    // 'true/false to print gridlines
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct PRINT_GRIDLINES_RECORD
    {
        public short opcode;
        public short length;
        public short PrintFlag;
    }

    // 'Integer record
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct tInteger
    {
        public short opcode;
        public short length;
        public short row;//unsigned integer
        public short col;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr1;//rgbAttr1 handles whether cell is hidden and/or locked
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr2;//rgbAttr2 handles the Font# and Formatting assigned to this cell
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr3;//rgbAttr3 handles the Cell Alignment/borders/shading
        public short intValue;//the actual integer value
    }

    // 'Number record
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct tNumber
    {
        public short opcode;
        public short length;
        public short row;
        public short col;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr1;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr2;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr3;
        public double NumberValue;//8 Bytes
    }
    //
    // 'Label (Text) record
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct tText
    {
        public short opcode;
        public short length;
        public short row;
        public short col;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr1;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr2;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr3;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte TextLength;
    }
    //
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct MARGIN_RECORD_LAYOUT
    {
        public short opcode;
        public short length;
        public double MarginValue;//8 bytes
    }
    //
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct HPAGE_BREAK_RECORD
    {
        public short opcode;
        public short length;
        public short NumPageBreaks;
    }
    //
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct DEF_ROWHEIGHT_RECORD
    {
        public int opcode;
        public int length;
        public int RowHeight;
    }
    //
    [StructLayoutAttribute(LayoutKind.Sequential, CharSet = CharSet.Auto, Pack = 1)]
    struct ROW_HEIGHT_RECORD
    {
        public int opcode;//08
        public int length;//should always be 16 bytes
        public int RowNumber;
        public int FirstColumn;
        public int LastColumn;
        public int RowHeight;//written to file as 1/20ths of a point
        public int internals;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte DefaultAttributes;//set to zero for no default attributes
        public int FileOffset;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr1;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr2;
        [MarshalAs(UnmanagedType.I1, SizeConst = 1)] public byte rgbAttr3;
    }

}