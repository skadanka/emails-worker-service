using Microsoft.EntityFrameworkCore;
using emails_worker_service.Models.FormModel;

public class ApplicationDbContext : DbContext
{
    public DbSet<FormModel> FormModels { get; set; }  // DBSet for FormModels

    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options)
    {
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // Specify the table name if different from the DbSet name
        modelBuilder.Entity<FormModel>()
            .ToTable("FormEntries");

        // Set primary key
        modelBuilder.Entity<FormModel>()
            .HasKey(fm => fm.MailId);

        // Configure fields with specific requirements or constraints
        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.TenderType)
            .IsRequired()
            .HasMaxLength(255);  // Assuming a max length for varchar fields

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.JobNumber)
            .IsRequired()
            .HasMaxLength(50);

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.SubmissionDate)
            .IsRequired();

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.SubmissionTime)
            .IsRequired();

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.FirstName)
            .IsRequired()
            .HasMaxLength(100);

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.LastName)
            .IsRequired()
            .HasMaxLength(100);

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.Phone)
            .IsRequired()
            .HasMaxLength(20); // Ensure the phone number fits into the field

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.Email)
            .IsRequired()
            .HasMaxLength(100);

        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.Exposure)
            .IsRequired()
            .HasMaxLength(255);

        // Optional: Define index for commonly queried fields
        modelBuilder.Entity<FormModel>()
            .HasIndex(fm => fm.Email);

        // Configure CvContent as a long text field
        modelBuilder.Entity<FormModel>()
            .Property(fm => fm.CvContent)
            .HasColumnType("text");  // Use "text" for databases like MySQL

        // Optional: Configure relationships if there are any related entities
        // For example, if there were an entity related to exposure types:
        //modelBuilder.Entity<FormModel>()
        //    .HasOne<ExposureType>()  // Assuming an entity named ExposureType
        //    .WithMany()
        //    .HasForeignKey(fm => fm.ExposureTypeId); // Assuming a foreign key named ExposureTypeId in FormModel
    }
}
