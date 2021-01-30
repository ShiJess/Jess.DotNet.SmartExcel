using System;

namespace Jess.SmartExcel
{
    /// <summary>
    /// 所有的枚举类型定义
    /// </summary>

    public enum ValueTypes
    {
        Integer = 0,
        Number = 1,
        Text = 2,
    }

    public enum CellAlignment
    {
        GeneralAlign = 0,
        LeftAlign = 1,
        CentreAlign = 2,
        RightAlign = 3,
        FillCell = 4,
        LeftBorder = 8,
        RightBorder = 16,
        TopBorder = 32,
        BottomBorder = 64,
        Shaded = 128
    }

    // 'used by rgbAttr2
    //'bits 0-5 handle the *picture* formatting, not bold/underline etc...
    //'bits 6-7 handle the font number
    public enum CellFont
    {
        Font0 = 0,
        Font1 = 64,
        Font2 = 128,
        Font3 = 192,
    }

    // 'used by rgbAttr1
    // 'bits 0-5 must be zero
    // 'bit 6 locked/unlocked
    // 'bit 7 hidden/not hidden
    public enum CellHiddenLocked
    {
        Normal = 0,
        Locked = 64,
        Hidden = 128,
    }

    public enum MarginTypes
    {
        LeftMargin = 38,
        RightMargin = 39,
        TopMargin = 40,
        BottomMargin = 41,
    }

    /// <summary>
    ///
    /// </summary>
    /// <remarks>可以使用|符号进行或组合</remarks>
    public enum FontFormatting
    {
        NoFormat = 0,
        Bold = 1,
        Italic = 2,
        Underline = 4,
        Strikeout = 8
    }
}