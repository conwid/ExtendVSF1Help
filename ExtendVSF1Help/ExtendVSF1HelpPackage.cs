using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace ExtendVSF1Help
{

    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists, PackageAutoLoadFlags.BackgroundLoad)]
    [Guid(ExtendVSF1HelpPackage.PackageGuidString)]
    public sealed class ExtendVSF1HelpPackage : AsyncPackage
    {

        public const string PackageGuidString = "e85ea583-38ae-4539-b2a5-037b5d52ce53";

        private CommandEvents commandEvents;
        private DTE2 dte;

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            commandEvents = dte.Events.CommandEvents[typeof(VSConstants.VSStd97CmdID).GUID.ToString("B"), (int)VSConstants.VSStd97CmdID.F1Help];
            commandEvents.AfterExecute += OnAfterExecute;
        }

        private void FetchContextAttributes(ContextAttributes contextAttributes, List<string> totalAttributes)
        {
            if (contextAttributes == null)
                return;

            contextAttributes.Refresh();
            foreach (ContextAttribute contextAttribute in contextAttributes)
            {
                var values = ((ICollection)contextAttribute.Values).Cast<string>();
                string attribute = $"{contextAttribute.Name}={string.Join(";", values)}";
                totalAttributes.Add(attribute);
            }
        }
        private void OnAfterExecute(string guid, int id, object customIn, object customOut)
        {
            Window activeWindow = dte.ActiveWindow;
            ContextAttributes contextAttributes = activeWindow.DTE.ContextAttributes;
            contextAttributes.Refresh();
            var attributes = new List<string>();

            try { FetchContextAttributes(contextAttributes?.HighPriorityAttributes, attributes); } catch { }
            try { FetchContextAttributes(contextAttributes, attributes); } catch { }

            HandleHelpCall(attributes);
        }

        private void HandleHelpCall(List<string> attributes)
        {
            var keywords = attributes.SingleOrDefault(a => a.StartsWith("keyword"));
            if (keywords is null)
                return;

            var keyword = keywords.Split(';').First().Split('=')[1];
            System.Diagnostics.Process.Start($"https://google.com/search?q={keyword}");
        }
    }
}
