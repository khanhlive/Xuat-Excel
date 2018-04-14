using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WebApplication1.ExcelCore
{
    public class ExportToExcelMultiple
    {
        private DataTable[] dataSource;
        private XLPaperSize[] paperSize;
        private bool[] tableHeaderBold;
        private XLPageOrientation[] pageOrientation;// = XLPageOrientation.Default;
        private XLBorderStyleValues[] boderStyle;
        private bool[] _wrapText;
        private string sheetsName;
        private bool[] _showGridLine;
        private ExcelColumnContent[] columnsWidth;
        private XLAlignmentHorizontalValues[] headerHorizontalAlignment;

        public XLPageOrientation[] PageOrientation
        {
            get { return pageOrientation; }
            set { pageOrientation = value; }
        }

        public XLPaperSize[] PaperSize
        {
            get { return paperSize; }
            set { paperSize = value; }
        }

        public bool[] TableHeaderBold
        {
            get { return tableHeaderBold; }
            set { tableHeaderBold = value; }
        }

        public ExcelColumnContent[] ColumnsWidth
        {
            get { return columnsWidth; }
            set { columnsWidth = value; }
        }

        public XLAlignmentHorizontalValues[] HeaderHorizontalAlignment
        {
            get { return headerHorizontalAlignment; }
            set { headerHorizontalAlignment = value; }
        }

        public DataTable[] DataSource
        {
            get { return dataSource; }
            set { dataSource = value; }
        }
        
        public XLBorderStyleValues[] BoderStyle
        {
            get { return boderStyle; }
            set { boderStyle = value; }
        }

        public bool[] ShowGridLine
        {
            get { return _showGridLine; }
            set { _showGridLine = value; }
        }

        public bool[] WrapText
        {
            get { return _wrapText; }
            set { _wrapText = value; }
        }
    }
}