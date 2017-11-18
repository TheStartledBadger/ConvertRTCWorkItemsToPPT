namespace CsvToPPT
{
    public class WorkItemInfo
    {
        public string Id { get; }
        public string Summary { get; }
        public string Owner { get; }
        public int Cost { get; }
        public string PlannedFor { get; }

        public WorkItemInfo(string id, string summary, string owner, int cost, string plannedFor)
        {
            this.Id = id;
            this.Summary = summary;
            this.Owner = owner;
            this.Cost = cost;
            this.PlannedFor = plannedFor;
        }
    }
}
