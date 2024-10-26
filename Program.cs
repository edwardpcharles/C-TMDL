using System; // Namespace for fundamental classes and base classes.
using System.IO; // Namespace for handling file and data streams.
using Microsoft.AnalysisServices; // Namespace for working with Analysis Services.
using Microsoft.AnalysisServices.Tabular.Extensions; // Namespace for tabular model extensions.

namespace powerbivideodemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Loop to continuously prompt user for options until exit.
            while (true)
            {
                Console.WriteLine("Choose an option:");
                Console.WriteLine("1. Export Semantic Model to TMDL");
                Console.WriteLine("2. Publish TMDL to Semantic Model");
                Console.WriteLine("3. Exit");
                string choice = Console.ReadLine() ?? string.Empty;;

                // Switch case to handle user's choice.
                switch (choice)
                {
                    case "1":
                        ExportModelToTmdl();
                        break;
                    case "2":
                        publishTmdl();
                        break;
                    case "3":
                        return; // Exit the program.
                    default:
                        Console.WriteLine("Invalid option. Please try again.");
                        break;
                }
            }
        }

        static void ExportModelToTmdl()
        {
            Console.WriteLine();
            Console.WriteLine("Enter Your XMLA Endpoint (Format: powerbi://api.powerbi.com/v1.0/myorg/Developer): ");
            var workspaceXmla = Console.ReadLine();
            //"powerbi://api.powerbi.com/v1.0/myorg/Developer";
            Console.WriteLine();
            Console.WriteLine("Enter Your Semantic Model Name: ");
            var datasetName = Console.ReadLine();
            //"Color Picker"
            Console.WriteLine();
            Console.WriteLine("Enter Path To Save TMDL Folder (Format: C:\\Users\\edwar\\OneDrive\\Desktop\\TMDL): ");
            var outputPath = Console.ReadLine();
            //"C:\Users\edwar\OneDrive\Desktop\TMDL";
            try
            {
                using (var server = new Server())
                {
                    server.Connect(workspaceXmla); // Connect to the server using XMLA endpoint.
                    var database = server.Databases.GetByName(datasetName); // Get the database by name.
                    var destinationFolder = $"{outputPath}\\{database.Name}-tmdl"; // Define the path to save TMDL folder.
                    TmdlSerializer.SerializeDatabaseToFolder(database, destinationFolder); // Serialize the database to TMDL folder.
                }
                Console.WriteLine();
                Console.WriteLine("Model Downloaded Successfully");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("An Error Occurred. Check your Inputs");
                Console.WriteLine($"Exception Message: {ex.Message}");
                Console.WriteLine();
            }
        }

        static void publishTmdl()
        {
            Console.WriteLine();
            Console.WriteLine("Enter Your XMLA Endpoint (Format: powerbi://api.powerbi.com/v1.0/myorg/Developer): ");
            var workspaceXmla = Console.ReadLine();
            //"powerbi://api.powerbi.com/v1.0/myorg/Developer";
            Console.WriteLine();
            Console.WriteLine("Enter Path To TMDL Folder To Publish (Format: C:\\Users\\edwar\\OneDrive\\Desktop\\TMDL): ");
            var tmdlFolderPath = Console.ReadLine();
            //"C:\Users\edwar\OneDrive\Desktop\TMDL\Color Picker-tmdl";
            try
            {
                var model = TmdlSerializer.DeserializeDatabaseFromFolder(tmdlFolderPath); // Deserialize the TMDL folder to a database model.
                using (var server = new Server())
                {
                    server.Connect(workspaceXmla); // Connect to the server using XMLA endpoint.
                    var existingDatabase = server.Databases.FindByName(model.Name); // Find the existing database by name.
                    if (existingDatabase != null)
                    {
                        model.Model.CopyTo(existingDatabase.Model); // Copy the model to the existing database.
                        existingDatabase.Model.SaveChanges(); // Save changes to the model.
                    }
                    else
                    {
                        var newDatabase = new Database
                        {
                            Name = model.Name,
                            ID = model.ID
                        };
                        server.Databases.Add(newDatabase); // Add the new database to the server.
                        newDatabase.Update(UpdateOptions.ExpandFull); // Update the new database.
                        model.Model.CopyTo(newDatabase.Model); // Copy the model to the new database.
                        newDatabase.Model.SaveChanges(); // Save changes to the model.
                    }
                }
                Console.WriteLine();
                Console.WriteLine("Model published successfully.");
                Console.WriteLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine("An Error Occurred. Check your Inputs or Your TMDL");
                Console.WriteLine($"Exception Message: {ex.Message}");
                Console.WriteLine();
            }
        }
    }
}
