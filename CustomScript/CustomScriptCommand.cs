using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Newtonsoft.Json;
using Task = System.Threading.Tasks.Task;

namespace CustomScript {
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CustomScriptCommand {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("1f7fe306-9d62-42fd-80cd-16c905b6e65e");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;


        private DTE2 dte2;
        private CustomScriptItemMenuCommand dynamicMenuCommand;
        /// <summary>
        /// Initializes a new instance of the <see cref="CustomScriptCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CustomScriptCommand(AsyncPackage package, OleMenuCommandService commandService) {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(this.Execute, menuCommandID);

            commandService.AddCommand(menuItem);

            dynamicMenuCommand = new CustomScriptItemMenuCommand(new CommandID(CommandSet, CommandId + 1),
              IsValidDynamicItem,
              OnInvokedDynamicItem,
              OnBeforeQueryStatusDynamicItem);

            commandService.AddCommand(dynamicMenuCommand);



            dte2 = (DTE2)this.ServiceProvider.GetServiceAsync(typeof(DTE)).GetAwaiter().GetResult();

        }

        private void OnInvokedDynamicItem(object sender, EventArgs args) {
            CustomScriptItemMenuCommand invokedCommand = (CustomScriptItemMenuCommand)sender;

            if (!File.Exists(invokedCommand.Path)) {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"找不到 {Path.GetFileName(invokedCommand.Path)} 檔案",
                    "找不到腳本路徑",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }

            var proc = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo() {
                FileName = invokedCommand.Path,
                WorkingDirectory = Path.GetDirectoryName(invokedCommand.Path),
            });
            proc.WaitForExit();

            if (proc.ExitCode == 0) {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "自訂腳本執行完成",
                    "執行成功",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            } else {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    "自訂腳本執行失敗",
                    "執行失敗",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return;
            }
        }

        private void OnBeforeQueryStatusDynamicItem(object sender, EventArgs args) {
            CustomScriptItemMenuCommand matchedCommand = (CustomScriptItemMenuCommand)sender;
            matchedCommand.Enabled = true;
            matchedCommand.Visible = true;


            var items = JsonConvert.DeserializeObject<CustomScriptItem[]>(File.ReadAllText(ConfigPath()));
            var index = matchedCommand.MatchedCommandId - 257;
            if (index < 0) index = 0;

            matchedCommand.Text = items[index].Name;
            matchedCommand.Path = items[index].Path;

            if (!Path.IsPathRooted(matchedCommand.Path)) {
                matchedCommand.Path = Path.Combine(GetProjectPath(), matchedCommand.Path);
            }

            matchedCommand.MatchedCommandId = 0;
        }

        private bool IsValidDynamicItem(int commandId) {
            if (!File.Exists(ConfigPath())) {
                return false;
            }

            var items = JsonConvert.DeserializeObject<CustomScriptItem[]>(File.ReadAllText(ConfigPath()));

            if ((commandId > (int)CommandId) &&
                ((commandId - (int)CommandId) <= items.Length)) {
                return true;
            } else {
                return false;
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CustomScriptCommand Instance {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider {
            get {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package) {
            // Switch to the main thread - the call to AddCommand in CustomScriptCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CustomScriptCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e) {
            ThreadHelper.ThrowIfNotOnUIThread();


            if (!File.Exists(ConfigPath())) {
                File.WriteAllText(ConfigPath(), "[{\"Name\":\"自訂腳本1\", \"Path\":\"test.bat\"}]");
            }

            dte2.ItemOperations.OpenFile(ConfigPath());
        }

        private string GetProjectPath() {
            IntPtr hierarchyPointer, selectionContainerPointer;
            Object selectedObject = null;
            IVsMultiItemSelect multiItemSelect;
            uint projectItemId;

            IVsMonitorSelection monitorSelection =
                    (IVsMonitorSelection)Package.GetGlobalService(
                    typeof(SVsShellMonitorSelection));

            monitorSelection.GetCurrentSelection(out hierarchyPointer,
                                                 out projectItemId,
                                                 out multiItemSelect,
                                                 out selectionContainerPointer);

            IVsHierarchy selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                                                 hierarchyPointer,
                                                 typeof(IVsHierarchy)) as IVsHierarchy;

            if (selectedHierarchy != null) {
                ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(
                                                  projectItemId,
                                                  (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                  out selectedObject));
            }

            Project selectedProject = selectedObject as Project;

            return Path.GetDirectoryName(selectedProject.FullName);
        }
        private string ConfigPath() {
            string projectPath = GetProjectPath();

            return Path.Combine(projectPath, "customScript.json");
        }
    }
}
