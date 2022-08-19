namespace Test;

[Table("TblPerson")]
public class Person
{
    [Key, Column(Order = 0)]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    [Display(Name = "Id", Description = "Inform an integer between 1 and 99999.")]
    [Range(1, 99999)]
    public int? Id { get; set; }

    [Display(Name = "Full Name", Description = "Name and Surname.")]
    [Required(ErrorMessage = "Complete name is required.")]
    [RegularExpression(@"^[a-zA-Z''-'\s]{1,40}$", ErrorMessage = "Numbers and Special Characteres are not allowed in Name")]
    public string? Name { get; set; } = "";

    [Required(ErrorMessage = "Email is Required")]
    [StringLength(100, MinimumLength = 5, ErrorMessage = "Email must have between 5 and 100 Characteres")]
    public string? Email { get; set; } = "";

    [DataType(DataType.DateTime)]
    public DateTime? Date { get; set; }

    [DataType(DataType.PhoneNumber)]
    public int? Phone { get; set; }
 

}
