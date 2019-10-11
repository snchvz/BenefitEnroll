using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewEnrollmentsProgram
{
    public class CompanyStatic
    {
        public static CompanyStatic Instance = new CompanyStatic();

        public string companyName;

        public CompanyStatic()
        {
            if (Instance == null)
            {
                Instance = this;
            }
        }
    }
}
