using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPO.ClientManager.Model
{
    public class SiteColumnProperty
    {
        public string DisplayName { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string Format { get; set; }
        public string Group { get; set; }
        public bool  IsRequired { get; set; }
        public string  DefaultValue { get; set; }
        public Guid  List { get; set; }
        public string  ShowField { get; set; }
        public string  ChoicesDetails { get; set; }
        public string NumLines { get; set; }
        public bool RichText { get; set; }
        public string RichTextMode { get; set; }
        public bool IsolateStyles { get; set; }
        public bool Sortable { get; set; }
        public string UserSelectionMode { get; set; }
        public Guid WebId { get; set; }

    }
}
