using System.ComponentModel;
using Inventor;

namespace SelectionInfo2
{
    /// <summary>
    /// This class extends <see cref="DocumentInfo"/> with information specific to Inventor.ComponentOccurrence objects.
    /// </summary>
    /// <seealso cref="SelectionInfo.DocumentInfo" />
    class OccurrenceInfo : DocumentInfo
    {
        private readonly ComponentOccurrence occurrence;

        /// <summary>
        /// Initializes a new instance of the <see cref="OccurrenceInfo"/> class.
        /// </summary>
        /// <param name="occurrence">The Inventor.ComponentOccurrence object.</param>
        public OccurrenceInfo(ComponentOccurrence occurrence) : base(occurrence.Definition.Document as Document)
        {
            this.occurrence = occurrence;
        }

        /// <summary>
        /// Gets or sets the display name of the occurrence.
        /// </summary>
        /// <value>
        /// The display name.
        /// </value>
        [Category("Occurrence")]
        public string DisplayName
        {
            get => occurrence.Name;
            set => occurrence.Name = value;
        }

        public string GetParentString(ComponentOccurrence occ)
        {
            if (occ == null)
                return "<null>";

            // Case 1: Parent is another occurrence (nested component)
            if (occ.ParentOccurrence != null)
            {
                return occ.ParentOccurrence.Name;
            }

            // Case 2: Parent is the top-level assembly
            if (occ.ParentOccurrence == null)
            {
                return "This is Top Level";
            }

            return "failure";
        }

        [Category("Occurrence")]
        public string ParentName
        {
            get => GetParentString(occurrence);
            //set => occurrence.Name = value;
        }
    }
}