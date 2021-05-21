using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace IcbcExcelMergeWeb.DotNetCore.Models
{

    // NOTE: Generated code may require at least .NET Framework 4.5 or .NET Core/Standard 2.0.
    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public partial class Reports
    {

        private ReportsReport reportField;

        /// <remarks/>
        public ReportsReport Report
        {
            get
            {
                return this.reportField;
            }
            set
            {
                this.reportField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class ReportsReport
    {

        private string nameField;

        private ReportsReportReportVal[] reportValField;

        /// <remarks/>
        public string Name
        {
            get
            {
                return this.nameField;
            }
            set
            {
                this.nameField = value;
            }
        }

        /// <remarks/>
        [System.Xml.Serialization.XmlElementAttribute("ReportVal")]
        public ReportsReportReportVal[] ReportVal
        {
            get
            {
                return this.reportValField;
            }
            set
            {
                this.reportValField = value;
            }
        }
    }

    /// <remarks/>
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public partial class ReportsReportReportVal
    {

        private byte reportRowField;

        private byte reportColField;

        private ushort valField;

        /// <remarks/>
        public byte ReportRow
        {
            get
            {
                return this.reportRowField;
            }
            set
            {
                this.reportRowField = value;
            }
        }

        /// <remarks/>
        public byte ReportCol
        {
            get
            {
                return this.reportColField;
            }
            set
            {
                this.reportColField = value;
            }
        }

        /// <remarks/>
        public ushort Val
        {
            get
            {
                return this.valField;
            }
            set
            {
                this.valField = value;
            }
        }
    }


}
