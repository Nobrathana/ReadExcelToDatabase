//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ReadExcelToDatabase.Entity
{
    using System;
    using System.Collections.Generic;
    
    public partial class tb_appraisal
    {
        public int id { get; set; }
        public Nullable<decimal> score { get; set; }
        public string description { get; set; }
        public Nullable<System.DateTime> monthAppraisal { get; set; }
        public Nullable<bool> status { get; set; }
        public System.DateTime created_at { get; set; }
        public string created_by { get; set; }
        public int empFK { get; set; }
        public Nullable<int> OT_FK { get; set; }
    
        public virtual tb_OT tb_OT { get; set; }
    }
}