using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Vidly.Dtos
{
    public class NewRentalDto
    {
        public int CustomerID { get; set; }

        public List<int> MovieIds { get; set; }
    }
}
