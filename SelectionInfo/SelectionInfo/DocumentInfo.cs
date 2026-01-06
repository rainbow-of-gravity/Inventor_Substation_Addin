using System.ComponentModel;
using Inventor;

namespace SelectionInfo2
{
    /// <summary>
    /// This class contains information about Document
    /// </summary>
    class DocumentInfo
    {
        private readonly Document document;
        private readonly DocumentiProperties iProperties;
        private readonly PhysicalProperties physicalProperties;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentInfo"/> class.
        /// </summary>
        /// <param name="document">An Inventor.Document object.</param>
        public DocumentInfo(Document document)
        {
            this.document = document;
            iProperties = new DocumentiProperties(document);
            physicalProperties = new PhysicalProperties(document);
        }

     /* /// <summary>
        /// Gets the model surface area in database units [cm2].
        /// </summary>
        /// <value>
        /// The area.
        /// </value>
        [Category("Physical")]
        public double Area => physicalProperties.Area();

        /// <summary>
        /// Gets the model mass in database units [kg].
        /// </summary>
        /// <value>
        /// The mass.
        /// </value>
        [Category("Physical")]
        public double Mass => physicalProperties.Mass();

        /// <summary>
        /// Gets the model mass in pounds [lb].
        /// </summary>
        /// <value>
        /// The mass.
        /// </value>
        [Category("Physical")]
        [DisplayName("Mass [lb]")]
        public double MassLb => physicalProperties.Mass("lb"); */

        /// <summary>
        /// Gets or sets the description iProperty value.
        /// </summary>
        /// <value>
        /// The Vendor.
        /// </value>
        [Category("iProperties")]
        public string Vendor
        {
            get => iProperties.Vendor;
            set => iProperties.Vendor = value;
        }

        /// <summary>
        /// Gets or sets the description iProperty value.
        /// </summary>
        /// <value>
        /// The Category.
        /// </value>
        [Category("iProperties")]
        public string Category
        {
            get => iProperties.Category;
            set => iProperties.Category = value;
            
        }


        /// <summary>
        /// Gets the part number iProperty value.
        /// </summary>
        /// <value>
        /// The part number.
        /// </value>
        [ReadOnly(true)]
        public string PartNumber
        {
            get => iProperties.PartNumber;
            set => iProperties.PartNumber = value;
        }

        /// <summary>
        /// Gets the name of the file.
        /// </summary>
        /// <value>
        /// The name of the file.
        /// </value>
        public string FileName => document.FullFileName;

        [Category("iProperties")]
        public string Ampacity
        {
            get => iProperties.UserDefined("Ampacity")?.ToString() ?? string.Empty;
            set => iProperties.UserDefined("Ampacity", value);
        }

        [Category("iProperties")]
        public string Mark
        {
            get => iProperties.UserDefined("Mark")?.ToString() ?? string.Empty;
            set => iProperties.UserDefined("Mark", value);
        }

        [Category("iProperties")]
        public string MFGNumber
        {
            get => iProperties.UserDefined("MFG Number")?.ToString() ?? string.Empty;
            set => iProperties.UserDefined("MFG Number", value);
            
        }

        [Category("iProperties")]
        public string OperatingNumber
        {
            get => iProperties.UserDefined("Operating Number")?.ToString() ?? string.Empty;
            set => iProperties.UserDefined("Operating Number", value);
        }

        [Category("iProperties")]
        public string VoltageClass
        {
            get => iProperties.UserDefined("Voltage Class")?.ToString() ?? string.Empty;
            set => iProperties.UserDefined("Voltage Class", value);
        }

        [Category("iProperties")]
        public string PowerRating
        {
            get => iProperties.UserDefined("Power Rating")?.ToString() ?? string.Empty;
            set => iProperties.UserDefined("Power Rating", value);
        }

    }
}