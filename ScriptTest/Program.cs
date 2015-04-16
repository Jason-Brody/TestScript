using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPTestScripts;
using System.Data;

namespace ScriptTest
{
    class Program
    {
        static void Main(string[] args)
        {

            RevaluationOfGLAccount script = new RevaluationOfGLAccount();
            
            script.Read(@"D:\test\SP01REVAL.txt");
            //script.GetReport("01.08.2014");
            
            
        }
    }
}
