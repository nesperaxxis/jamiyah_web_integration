using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Jamiyah_Web_Integration.Models
{
    // Models returned by MeController actions.
    public class GetViewModel
    {
        public string Hometown { get; set; }
    }
}