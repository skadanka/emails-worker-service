using emails_worker_service.Models.FormModel;

namespace emails_worker_service.Models
{
    public static class FormModelFactory
    {
        public static FormModelBase CreateFormModel(string formType)
        {
            switch (formType)
            {
                case "LinkedIn":
                    return new FormModelLinkedIn();
                case "Drushim":
                    return new FormModelDrushim();
                // Add more form models as needed
                default:
                    throw new ArgumentException("Invalid form type");
            }
        }
    }

}
