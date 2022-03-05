using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace AutoAnnotate
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Annotate
    {
        public const int CommandId = 4129;

        public static readonly Guid CommandSet = new Guid("747122d8-5365-4c08-900a-393c18cb8798");

        private readonly AsyncPackage package;

        private readonly DTE dteService;

        private Annotate(AsyncPackage package, OleMenuCommandService commandService, DTE dteService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));
            this.dteService = dteService ?? throw new ArgumentNullException(nameof(dteService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static Annotate Instance
        {
            get;
            private set;
        }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Annotate's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            DTE dteService = await package.GetServiceAsync(typeof(DTE)) as DTE;
            Instance = new Annotate(package, commandService, dteService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var language = this.dteService.ActiveDocument.ActiveWindow.Document.Language;
            
            if (language != Constants.CSharp)
            {
                return;
            }

            var fileCodeModel = dteService.ActiveDocument.ActiveWindow.Document.ProjectItem.FileCodeModel;

            PerformWithinUndoContext(() =>
            {
                AddUsingStatement(fileCodeModel);
                
                var firstCodeClass = GetFirstCodeClass(fileCodeModel);
                AddDataContractToCodeClass(firstCodeClass);
                AddDataMemberToPublicProperties(firstCodeClass);
            });
        }

        private CodeClass GetFirstCodeClass(FileCodeModel fileCodeModel)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var outermostElement in fileCodeModel.CodeElements)
            {
                if (outermostElement is CodeNamespace codeNamespace)
                {
                    foreach (var childOfNamespace in codeNamespace.Children)
                    {
                        if (childOfNamespace is CodeClass codeClass)
                        {
                            return codeClass;
                        }
                    }
                }
            }

            return null;
        }

        private void AddDataContractToCodeClass(CodeClass codeClass)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var item in codeClass.Attributes)
            {
                if (item is CodeAttribute existingCodeAttribute
                    && existingCodeAttribute.Name == Constants.DataContract)
                {
                    return;
                }
            }

            var codeAttribute = codeClass.AddAttribute(Constants.DataContract, null);
            RemoveParenthesesFromAttribute(codeAttribute);
        }

        private void AddDataMemberToPublicProperties(CodeClass codeClass)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var codeClassChild in codeClass.Children)
            {
                if (codeClassChild is CodeProperty codeProperty
                    && codeProperty.Access == vsCMAccess.vsCMAccessPublic)
                {
                    if (!IsDataMemberAlreadyExists(codeProperty))
                    {
                        var codeAttribute = codeProperty.AddAttribute(Constants.DataMember, null);
                        RemoveParenthesesFromAttribute(codeAttribute);
                    }
                }
            }
        }

        private bool IsDataMemberAlreadyExists(CodeProperty codeProperty)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var codePropertyChild in codeProperty.Children)
            {
                if (codePropertyChild is CodeAttribute codeAttribute
                    && codeAttribute.Name == Constants.DataMember)
                {
                    return true;
                }
            }

            return false;
        }

        private void RemoveParenthesesFromAttribute(CodeAttribute targetAttribute)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var startPoint = targetAttribute.GetStartPoint().CreateEditPoint();
            var endPoint = targetAttribute.GetEndPoint().CreateEditPoint();

            var nameWithoutParentheses = targetAttribute.Name;
            startPoint.ReplaceText(endPoint, nameWithoutParentheses, (int)vsEPReplaceTextOptions.vsEPReplaceTextAutoformat);
        }

        private void PerformWithinUndoContext(Action action)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // The UndoContext may have been opened by some other operations outside of this extension
            // Therefore we do not attempt to open / close it ourselves
            var isOpenedByThisExtension = false;

            try
            {
                if (dteService.UndoContext.IsOpen)
                {
                    isOpenedByThisExtension = false;
                }
                else
                {
                    dteService.UndoContext.Open("Add DataContract & DataMember");
                    isOpenedByThisExtension = true;
                }

                action();
            }
            catch (COMException comException)
            {
                // TODO add logging
            }
            finally
            {
                if (isOpenedByThisExtension)
                {
                    dteService.UndoContext.Close();
                }
            }
        }

        private void AddUsingStatement(FileCodeModel fileCodeModel)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var usingStatement = $"using {Constants.SystemRuntimeSerialization};";

            CodeNamespace foundCodeNamespace = null;
            CodeImport foundCodeImport = null;
            foreach (CodeElement element in fileCodeModel.CodeElements)
            {
                if (element is CodeNamespace codeNamespace)
                {
                    foundCodeNamespace = codeNamespace;
                }

                if (element is CodeImport codeImport)
                {
                    var importNamespace = codeImport.Namespace;

                    if (!importNamespace.StartsWith(Constants.System))
                    {
                        continue;
                    }

                    foundCodeImport = codeImport;

                    var stringCompareResult = StringComparer.InvariantCulture.Compare(Constants.SystemRuntimeSerialization, importNamespace);

                    if (stringCompareResult == 0)
                    {
                        return;
                    }

                    if (stringCompareResult < 0)
                    {
                        codeImport.GetStartPoint().CreateEditPoint()
                            .Insert($"{usingStatement}{Environment.NewLine}");
                        return;
                    }
                }
            }

            // If it reaches here, it means we have gone through all using statements
            // In that case, insert using statement as last position
            if (foundCodeImport != null)
            {
                foundCodeImport.GetEndPoint().CreateEditPoint()
                    .Insert($"{Environment.NewLine}{usingStatement}");
                return;
            }

            // If it reaches here, it means this file has no using statement
            // In that case, insert using statement before namespace declaration
            if (foundCodeNamespace != null)
            {
                foundCodeNamespace.GetStartPoint().CreateEditPoint()
                    .Insert($"{usingStatement}{Environment.NewLine}{Environment.NewLine}");
            }
        }
    }
}
