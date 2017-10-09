### SCTags

Add Tags to story

### How to install from Visual Studio

Create a new C# console application with the name SCTags.

In the project folder, replace program.cs and app.config with the ones from this repo.

Add References

Project -> Add References

System.Configuration


Install Packages

Tools -> Nuget Package Manager -> Package Manager Console 

Enter these lines in the console in this order.

Install-Package Newtonsoft.Json

Install-Package SharpCloud.ClientAPI

Install-Package EPPlus


Enter your Sharpcloud username, password,team, sheet in the app.config file.

### Issues: 
Downloads all images and references from the story. Might take some time to finish.

Program Crashes if you have any of the spreadsheets created from this program open during runtime.