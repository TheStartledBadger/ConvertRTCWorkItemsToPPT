using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CsvToPPT
{
    class WorkItemInfo
    {
        public string Id { get; }
        public string Summary { get; }
        public string Owner { get; }
        public int Cost { get; }

        public WorkItemInfo(string id, string summary, string owner, int cost)
        {
            this.Id = id;
            this.Summary = summary;
            this.Owner = owner;
            this.Cost = cost;
        }
    }
}
