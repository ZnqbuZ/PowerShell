using System.Management.Automation;
using System.Reflection;

namespace ZnqbuZ.PowerShell
{
    [Cmdlet(VerbsDiagnostic.Test, "Interactive")]
    [OutputType(typeof(bool))]
    public class TestInteractiveCmdlet : Cmdlet
    {
        private string[] InteractiveModes { get; set; } = ["NoExit", "Interactive"];
        private string[] NonInteractiveModes { get; set; } = ["Command", "File", "NonInteractive", "Version"];

        private void Log(string message)
        {
            WriteVerbose(message);
        }

        private bool IsInteractive(string[] args)
        {
            Log($"Received parameters: {string.Join(", ", args)}");

            var cppAssembly = Assembly.Load("Microsoft.PowerShell.ConsoleHost");
            var cppType = cppAssembly
                .GetType("Microsoft.PowerShell.CommandLineParameterParser")!;
            var cppMethods = cppType
                .GetMethods(BindingFlags.NonPublic | BindingFlags.Instance)
                .ToDictionary(m => m.Name);
            var cppProperties = cppType
                .GetProperties(BindingFlags.NonPublic | BindingFlags.Instance)
                .ToDictionary(p => p.Name);
            var cppFields = cppType
                .GetFields(BindingFlags.NonPublic | BindingFlags.Instance)
                .ToDictionary(f => f.Name);

            var cppConstructorInfo = cppType.GetConstructor(
                BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null)!;
            var cppObject = cppConstructorInfo.Invoke(null);

            cppMethods["Parse"].Invoke(cppObject, [args]);

            if ((bool)(cppProperties["AbortStartup"].GetValue(cppObject)!))
            {
                Log("Invalid parameters.");
                return false;
            }

            if ((bool)cppProperties["NoPrompt"].GetValue(cppObject)!)
            {
                return false;
            }

            if (InteractiveModes.Contains("NoExit")
                && (bool)cppProperties["NoExit"].GetValue(cppObject)!)
            {
                return true;
            }

            if (NonInteractiveModes.Contains("Command")
                && (string?)cppProperties["InitialCommand"].GetValue(cppObject) != null)
            {
                return false;
            }

            if (NonInteractiveModes.Contains("File")
                && (string?)cppProperties["File"].GetValue(cppObject) != null)
            {
                return false;
            }

            if ((NonInteractiveModes.Contains("NonInteractive") || InteractiveModes.Contains("Interactive"))
                && (bool)cppProperties["NonInteractive"].GetValue(cppObject)!)
            {
                return false;
            }

            if ((NonInteractiveModes.Contains("Version"))
                && (bool)cppProperties["ShowVersion"].GetValue(cppObject)!)
            {
                return false;
            }

            return true;
        }

        protected override void ProcessRecord()
        {
            WriteObject(IsInteractive(Environment.GetCommandLineArgs().Skip(1).ToArray()));
        }
    }
}