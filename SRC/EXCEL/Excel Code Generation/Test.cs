using System;

[DisplayTable("Test")]
public class Test
{
    [DisplayColumn("Custom Property First Name", 0)]
    public System.String CustomPropertyFirstName { get; set; }
    [DisplayColumn("Custom Property Last Name", 1)]
    public System.String CustomPropertyLastName { get; set; }
    [DisplayColumn("Custom Property Age", 2)]
    public System.String CustomPropertyAge { get; set; }
}
