namespace akelon_1.Classes;

public class Request
{
    public int Code { get; set; }
    public int ProductCode { get; set; }
    public int ClientCode { get; set; }
    public int Number { get; set; }
    public int Count { get; set; }
    public DateOnly PlacementDate { get; set; }
}