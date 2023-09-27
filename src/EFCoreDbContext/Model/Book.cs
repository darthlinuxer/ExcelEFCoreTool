using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using OfficeOpenXml.Attributes;

namespace Model;

[Table("TblBooks")]
public class Book
{
    [Key, Column(Order = 0)]
    [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
    [Display(Name = "Id", Description = "Inform an integer between 1 and 99999.")]
    [Range(1, 99999)]
    public int? Id { get; set; }

    [Display(Name = "Book Name")]
    [Required(ErrorMessage = "Complete name is required.")]
    [RegularExpression(@"^[a-zA-Z''-'\s]{1,40}$", ErrorMessage = "Numbers and Special Characteres are not allowed in Name")]
    public string? Name { get; set; } = "";

    [EpplusIgnore]
    [ForeignKey("PersonFK")]
    public Person? Person { get; set; }
    public int? PersonFK { get; set; }


}
