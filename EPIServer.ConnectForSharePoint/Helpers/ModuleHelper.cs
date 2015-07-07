using System.Collections.Generic;
using System.Reflection;
using EPiServer.DataAbstraction;
using EPiServer.Shell;
using EPiServer.Shell.Modules;
using EPiServer.Web.Hosting;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Helper class to provide utilities to resolve Url to the module
    /// </summary>
    public static class ModuleHelper
    {
        public const string PublicVirtualPath = "/util/SharePoint/";

        /// <summary>
        /// Readonly propery for geting the module path.
        /// </summary>
        public static string MyModulePath
        {
            get
            {
                return Paths.ToResource(typeof(ModuleHelper), string.Empty);
            }
        }

        /// <summary>
        /// Return language ID of the top enabled language
        /// </summary>
        /// <returns>A string of language's ID</returns>
        public static string DefaultLanguageID
        {
            get
            {
                return LanguageBranch.LanguageBranchRepository.LoadFirstEnabledBranch().LanguageID;
            }
        }

        /// <summary>
        /// <example>
        /// Call ModuleHelper.ToClientResource("X") will return "/secure/EPiServer.MODULENAME/1.0.0.0/X" where 1.0.0.0 is the clientResourceRelativePath in module.config
        /// </example>
        /// </summary>
        /// <param name="virtualPath"></param>
        /// <returns></returns>
        public static string ToClientResource(string virtualPath)
        {
            return Paths.ToClientResource(typeof(ModuleHelper), virtualPath);
        }

        public static string ToResource(string virtualPath)
        {
            return Paths.ToResource(typeof(ModuleHelper), virtualPath);
        }

        public static string ToPhysicalVPP(string virtualpath)
        {
            var ret = string.Empty;
            var vpp = VirtualPathHandler.GetProvider("ProtectedAddons") as VirtualPathNonUnifiedProvider;
            if (vpp == null)
            {
                return ret;
            }

            // now we have vpp.LocalPath point to C:\EPiServer\VPP\MyEPiServerSite\Modules\
            ret = string.Format(vpp.LocalPath + @"\{0}", virtualpath).Replace('/', '\\');
            return ret;
        }

        /// <summary>
        /// <example>Output could be "C:\\EPiServer\\VPP\\MyEPiServerSite\\Modules\\EPiServer.MODULENAME\\virtualpath"</example>
        /// </summary>
        /// <param name="virtualpath"></param>
        /// <returns></returns>
        public static string ToPhysicalVPPResource(string virtualpath)
        {
            // pathToModuleResource will be "EPiServer.MODULENAME\", we use this technique to get correct MODULENAME and VersionString in VPP physical path
            var pathToModuleResource = ModuleHelper.ToResource(virtualpath).Replace(Paths.ProtectedRootPath, "");
            return ToPhysicalVPP(pathToModuleResource);
        }

        /// <summary>
        /// <example>Output could be "C:\\EPiServer\\VPP\\MyEPiServerSite\\Modules\\EPiServer.MODULENAME\\1.0.0.0\\virtualpath"</example>
        /// </summary>
        /// <param name="virtualpath"></param>
        /// <returns></returns>
        public static string ToPhysicalVPPClientResource(string virtualpath)
        {
            // pathToModuleClientResource will be "EPiServer.MODULENAME\[VERSION]\"
            var pathToModuleClientResource = ModuleHelper.ToClientResource(virtualpath).Replace(Paths.ProtectedRootPath, "");
            return ToPhysicalVPP(pathToModuleClientResource);
        }

        /// <summary>
        /// Return a list of ebabled languages on the site
        /// </summary>
        /// <returns> A list of <see cref="LanguageBranch"/></returns>
        public static IList<LanguageBranch> GetEnabledLanguages()
        {
            return LanguageBranch.LanguageBranchRepository.ListEnabled();
        }

        /// <summary>
        /// Get current ShellModule (describe in module.config) of this addon
        /// </summary>
        /// <returns></returns>
        public static ShellModule GetShellModule()
        {
            var moduleTable = ServiceLocation.ServiceLocator.Current.GetInstance<ModuleTable>();
            return moduleTable.GetModuleOrDefault(Assembly.GetExecutingAssembly());
        }
    }
}