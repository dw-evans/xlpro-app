using ExcelDna.Integration;

public static class MyFunctions
{
    [ExcelFunction(Description = "My first .NET function")]
    public static string SayHello(string name)
    {
        return "Hello " + name;
    }

    [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
    public static double AddThem(
        [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] 
        double v1,
        [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     
        double v2)
    {
        return v1 + v2;
    }
}