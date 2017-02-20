using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using EnvDTE80;

namespace UnlockDll
{
    /// <summary>
    /// Class to find paths of all blocking dlls and unblock it.
    /// </summary>
    public class DllUnblocker
    {
        #region Private Fields
        private const string PatternForPath = @"'[^']+'";
        private const string StringToSearch = "Access to the path";
        private const string SuccessReadOnlyMessage = "Success - File : {0}, Removed ReadOnly attribute.\n";
        private const string SuccessHiddenMessage = "Success - File : {0}, Removed Hidden attribute.\n";
        private const string ErrorFileNotFoundMessage = "Error - File : {0} does not exist ! \n";
        private EnvDTE80.DTE2 envDte2;
        #endregion

        #region Constructors and Public Methods
        /// <summary>
        /// Initializes a new instance of the <see cref="DllUnblocker"/> class.
        /// </summary>
        /// <param name="dteObject">The DTE object.</param>
        public DllUnblocker(EnvDTE80.DTE2 dteObject)
        {
            envDte2 = dteObject;
        }

        /// <summary>
        /// Unblocks all DLLS.
        /// </summary>
        /// <returns></returns>
        public string UnblockAllDlls()
        {
            string result = string.Empty;
            var paths = GetPathsToUnlock();
            foreach (var path in paths)
            {
                result += RemoveBlockingAttributes(path);
            }
            return result;
        }

        #endregion

        #region  Private Methods
        private string PathToSolution => System.IO.Path.GetDirectoryName(envDte2.Solution.FullName);

        private List<ErrorItem> GetErrors()
        {
            ErrorItems errors = envDte2.ToolWindows.ErrorList.ErrorItems;
            List<ErrorItem> listOfSelecteedItems = new List<ErrorItem>();

            for (int i = 1; i <= errors.Count; i++)
            {
                ErrorItem item = errors.Item(i);
                listOfSelecteedItems.Add(item);
            }
            return listOfSelecteedItems;
        }

        private List<string> GetPathsToUnlock()
        {
            var errors = GetErrors();
            List<string> listOfSelecteedItems = new List<string>();
            string pathToSolution = PathToSolution;

            foreach (var error in errors)
            {
                string decription = error.Description;
                if (decription.Contains(StringToSearch))
                {
                    var dllPath = PathToDllFromError(decription).Replace("'", "");
                    if (System.IO.Path.IsPathRooted(dllPath))
                    {
                        listOfSelecteedItems.Add(dllPath);
                    }
                    else
                    {
                        var dllName = dllPath.Split(Path.DirectorySeparatorChar).Last();
                        var solutionDirectroy = new DirectoryInfo(pathToSolution);
                        FileInfo[] files = solutionDirectroy.GetFiles(dllName, SearchOption.AllDirectories);
                        files.ToList().ForEach(file => listOfSelecteedItems.Add(file.FullName));
                    }
                }
            }
            return listOfSelecteedItems;
        }

        private string PathToDllFromError(string errorMessage)
        {
            Regex r = new Regex(PatternForPath, RegexOptions.IgnoreCase);
            Match m = r.Match(errorMessage);
            string path = string.Empty;
            if (m.Success)
            {
                path = m.Value;
            }
            return path;
        }

        private string RemoveBlockingAttributes(string pathToDll)
        {
            string result = string.Empty;
            if (!File.Exists(pathToDll))
            {
                result = string.Format(ErrorFileNotFoundMessage, pathToDll);
                return result;
            }

            FileAttributes attributes = File.GetAttributes(pathToDll);

            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                attributes = RemoveAttribute(attributes, FileAttributes.ReadOnly);
                File.SetAttributes(pathToDll, attributes);
                result += string.Format(SuccessReadOnlyMessage, pathToDll);
            }

            if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
            {
                attributes = RemoveAttribute(attributes, FileAttributes.Hidden);
                File.SetAttributes(pathToDll, attributes);
                result += string.Format(SuccessHiddenMessage, pathToDll);
            }

            return result;
        }

        private static FileAttributes RemoveAttribute(FileAttributes attributes, FileAttributes attributesToRemove)
        {
            return attributes & ~attributesToRemove;
        }

        #endregion
    }
}