using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Entity;

namespace График_ПЗ
{
    class AppContext : DbContext
    {
        public DbSet<Employee> Employees { get; set; }
        
        public AppContext() : base("DefaultConnection")
        {
        }
    }
}
