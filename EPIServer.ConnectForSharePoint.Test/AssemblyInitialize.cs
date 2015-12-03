using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Test
{
    /// <summary>
    /// Summary description for AssemblyInitialize
    /// </summary>
    [TestClass]
    public class AssemblyInitialize
    {
        //testing build 2
        public static TestScopeHelper AssemblyScope { get; set; }

        [AssemblyInitialize]
        public static void Initialize(TestContext context)
        {
            // Run EPiServer test setup code.
            AssemblyScope = new TestScopeHelper();
            AssemblyScope.SetupCommonServices();
        }

        [AssemblyCleanup]
        public static void Cleanup()
        {
            AssemblyScope.Dispose();
            AssemblyScope = null;
        }
    }
}
