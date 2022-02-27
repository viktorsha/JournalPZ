using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Журнал_ПЗ
{
    public abstract class EmployeeEvent
    {
        protected int Id { get; set; }
        protected string Name { get; set; }
        protected string OrderProps { get; set; }
    }
}
