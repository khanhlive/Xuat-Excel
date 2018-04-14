using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Web;
using XuatExcelClosedXML.ExcelCore;

namespace XuatExcelClosedXML.Design
{
    
    public class WorksheetConfig : ICloneable
    {
        public WorksheetConfig()
        {
            this.Header = new HeaderComponent();
            this.Title = new  TitleComponent();
            this.TableHeader = new TableHeaderComponent();
            this.Content = new ContentComponent();
            this.Signature = new SignatureComponent();
            this.PageSetup = XLWorkbook.DefaultPageOptions;
            this.Style = new ExcelStyle() ;
            this.DataSource = new DataTable();
        }
        
        public string Name { get; set; }
        public HeaderComponent Header { get; set; }
        public TitleComponent Title { get; set; }
        public TableHeaderComponent TableHeader { get; set; }
        public ContentComponent Content { get; set; }
        public SignatureComponent Signature { get; set; }
        public IXLPageSetup PageSetup { get; set; }
        public ExcelStyle Style { get; set; }
        public bool ShowGridLine { get; set; }
        public DataTable DataSource { get; set; }
        public WorksheetConfig Clone(WorksheetConfig setting)
        {
            var serialized = JsonConvert.SerializeObject(setting);
            return JsonConvert.DeserializeObject<WorksheetConfig>(serialized);
        }


        object ICloneable.Clone()
        {
            throw new NotImplementedException();
        }
    }
    
    public class ExcelStyle : ICloneable
    {
        public ExcelStyle()
        {
            this.BackgroundColor = Color.White;
            this.FontColor = Color.Black;
            this.FontName = "Times New Roman";
            this.FontSize = 12;
            this.Underline = XLFontUnderlineValues.None;
            this.BottomBorder = XLBorderStyleValues.None;
            this.LeftBorder = XLBorderStyleValues.None;
            this.RightBorder = XLBorderStyleValues.None;
            this.TopBorder = XLBorderStyleValues.None;
        }
        public XLAlignmentHorizontalValues Horizontal { get; set; }
        public XLAlignmentVerticalValues Vertical { get; set; }
        public XLBorderStyleValues BottomBorder { get; set; }
        public XLBorderStyleValues LeftBorder { get; set; }
        public XLBorderStyleValues RightBorder { get; set; }
        public XLBorderStyleValues TopBorder { get; set; }
        public Color BackgroundColor { get; set; }
        public bool Bold { get; set; }
        public Color FontColor { get; set; }
        public string FontName { get; set; }
        public double FontSize { get; set; }
        public bool Italic { get; set; }
        public bool Strikethrough { get; set; }
        public XLFontUnderlineValues Underline { get; set; }

        public IXLStyle GetStyle(IXLStyle Style)
        {
            Style.Alignment.Horizontal = this.Horizontal;
            Style.Alignment.Vertical = this.Vertical;
            Style.Border.BottomBorder = this.BottomBorder;
            Style.Border.TopBorder = this.TopBorder;
            Style.Border.LeftBorder = this.LeftBorder;
            Style.Border.RightBorder = this.RightBorder;
            Style.Fill.BackgroundColor = XLColor.FromColor(this.BackgroundColor);
            Style.Font.Bold = this.Bold;
            Style.Font.FontColor = XLColor.FromColor(this.FontColor);
            Style.Font.FontName = this.FontName;
            Style.Font.FontSize = this.FontSize;
            Style.Font.Italic = this.Italic;
            Style.Font.Strikethrough = this.Strikethrough;
            Style.Font.Underline = this.Underline;
            //Style.Font.VerticalAlignment = this.VerticalAlignment;
            return Style;
        }

        public ExcelStyle Clone()
        {
            return (ExcelStyle)this.MemberwiseClone();
        }
        object ICloneable.Clone()
        {
            return Clone();
        }

    }
    public class HeaderComponent : ICloneable
    {
        public HeaderComponent()
        {
            this.HeaderLeft = new List<TextComponent>();
            this.HeaderRight = new List<TextComponent>();
            this.SubHeaderRight = new List<TextComponent>();
            this.SubHeaderLeft = new List<TextComponent>();
        }
        public List<TextComponent> HeaderLeft { get; set; }
        public List<TextComponent> HeaderRight { get; set; }
        public List<TextComponent> SubHeaderRight { get; set; }
        public List<TextComponent> SubHeaderLeft { get; set; }
        public HeaderComponent Clone()
        {
            return (HeaderComponent)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }

    public class TableHeaderComponent : ICloneable
    {
        public TableHeaderComponent()
        {
            Style = new ExcelStyle();
        }
        public ExcelStyle Style { get; set; }
        public ExcelColumnContent[] ColumnsWidth { get; set; }
        public TableHeaderComponent Clone()
        {
            return (TableHeaderComponent)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }

    public class ContentComponent : ICloneable
    {
        public ContentComponent()
        {
            BorderStyle =  XLBorderStyleValues.Thin;
        }
        public XLBorderStyleValues BorderStyle { get; set; }
        public bool WrapText { get; set; }
        public ContentComponent Clone()
        {
            return (ContentComponent)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }
    public class SignatureComponent : ICloneable
    {
        public SignatureComponent()
        {
            this.SignatureLeft = new List<TextComponent>();
            this.SignatureRight = new List<TextComponent>();
        }
        public List<TextComponent> SignatureLeft { get; set; }
        public List<TextComponent> SignatureRight { get; set; }
        public SignatureComponent Clone()
        {
            return (SignatureComponent)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }
    public class TitleComponent : ICloneable
    {
        public TitleComponent()
        {
            this.Title = new TextComponent();
            this.SubTitle = new TextComponent();
        }
        public TextComponent Title { get; set; }
        public TextComponent SubTitle { get; set; }
        public TitleComponent Clone()
        {
            return (TitleComponent)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }
    public class TextComponent : ICloneable
    {
        public TextComponent()
        {
            this.Style = new ExcelStyle() ;
        }
        public string Text { get; set; }
        public ExcelStyle Style { get; set; }

        public TextComponent Clone()
        {
            return (TextComponent)this.MemberwiseClone();
        }

        object ICloneable.Clone()
        {
            return Clone();
        }
    }
    public class ExcelCell
    {
        public ExcelCell()
        {
            this.Style = new ExcelStyle();
        }
        public ExcelCell(IXLCell cell, ExcelStyle style)
        {
            this.Style = style;
            this.Cell = cell;
        }
        public IXLCell Cell { get; set; }
        public ExcelStyle Style { get; set; }
    }
}