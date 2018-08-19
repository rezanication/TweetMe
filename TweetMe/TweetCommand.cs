using EnvDTE;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace TweetMe
{
    internal sealed class TweetCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("092c04e1-da08-4830-b0a7-47aaeb808e7a");

        private readonly Package package;

        private TweetCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static TweetCommand Instance
        {
            get;
            private set;
        }


        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new TweetCommand(package, commandService);
        }

        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DTE dte = (DTE)this.ServiceProvider.GetService(typeof(DTE));
            Assumes.Present(dte);
            Document activeDocument = dte.ActiveDocument;

            try
            {
                if (activeDocument != null && activeDocument?.Type == "Text")
                {
                    TextDocument text = (TextDocument)activeDocument.Object(String.Empty);
                    if (!text.Selection.IsEmpty)
                    {
                        if (text.Selection.Text.Length > 140)
                        {
                            VsShellUtilities.ShowMessageBox(ServiceProvider, "TweetMe: maximum length is 140 characters", "Oops!", OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                        }
                        else
                        {
                            var message = $"{text.Selection.Text}";
                            System.Diagnostics.Process.Start("https://twitter.com/intent/tweet?text=" + message);
                        }
                    }
                    else
                    {
                        VsShellUtilities.ShowMessageBox(ServiceProvider, "You need to select some code!", "Oops!", OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    }
                }
            }
            catch
            {
                VsShellUtilities.ShowMessageBox(ServiceProvider, "Something went wrong, try again!", "Oops!", OLEMSGICON.OLEMSGICON_INFO, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

        }

    }
}
