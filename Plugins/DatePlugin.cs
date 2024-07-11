using Microsoft.SemanticKernel;
using System.ComponentModel;

namespace BMCGen
{
    public class DatePlugin
    {
        [KernelFunction, Description("Get the current date and time")]
        public static string GetCurrentDateTime()
        {            
            return DateTime.Now.ToString();
        }
    }
}
