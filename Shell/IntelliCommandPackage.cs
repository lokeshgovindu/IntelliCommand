// -----------------------------------------------------------------------------
// License: Microsoft Public License (Ms-PL)
// -----------------------------------------------------------------------------

using EnvDTE80;
using IntelliCommand.Presentation;
using IntelliCommand.Services;
using IntelliCommand.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using Task = System.Threading.Tasks.Task;

namespace IntelliCommand
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(IntelliCommandPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideOptionPage(typeof(IntelliCommandOptionsDialogPage), "IntelliCommand", "General", 0, 0, true)]
    public sealed class IntelliCommandPackage : AsyncPackage, IAppServiceProvider
    {
        /// <summary>
        /// IntelliCommandPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "9aecd086-965c-40df-8ae5-db4e0ad4db4b";

        private static IAppServiceProvider appServiceProvider;

        private readonly Dictionary<Type, object> registeredServices = new Dictionary<Type, object>();

        private CommandsInfoWindow mainWindow;

        /// <inheritdoc />
        public TService GetService<TService>()
        {
            return this.GetService<TService, TService>();
        }

        /// <inheritdoc />
        public TServiceImpl GetService<TService, TServiceImpl>()
        {
            object service;
            if (this.registeredServices.TryGetValue(typeof(TService), out service))
            {
                Debug.Assert(service is TService, "service is TService");
                return (TServiceImpl)service;
            }

            return (TServiceImpl)this.GetService(typeof(TService));
        }

        internal static IAppServiceProvider GetAppServiceProvider()
        {
            return IntelliCommandPackage.appServiceProvider;
        }

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            var shellService = this.GetService(typeof(SVsShell)) as IVsShell;
            if (shellService != null)
            {
                InitializeCommandServices();
            }
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            if (disposing)
            {
                foreach (var registeredService in this.registeredServices.Values)
                {
                    var disposable = registeredService as IDisposable;
                    if (disposable != null)
                    {
                        disposable.Dispose();
                    }
                }

                if (this.mainWindow != null)
                {
                    this.mainWindow.Close();
                }
            }
        }

        private void InitializeCommandServices()
        {
            var dte = this.GetService<SDTE>() as DTE2;
            var outputWindowService = new OutputWindowService(this);
            this.registeredServices.Add(typeof(IOutputWindowService), outputWindowService);

            Debug.Assert(dte != null, "dte != null");
            if (dte != null)
            {
                this.registeredServices.Add(typeof(ICommandScopeService), new CommandScopeService(this));
                this.registeredServices.Add(typeof(Dispatcher), Dispatcher.CurrentDispatcher);
                this.registeredServices.Add(typeof(IKeyboardListenerService), new KeyboardListenerService());
                this.registeredServices.Add(typeof(IPackageSettings), this.GetDialogPage(typeof(IntelliCommandOptionsDialogPage)));

                IntelliCommandPackage.appServiceProvider = this;

                this.mainWindow = new CommandsInfoWindow(this) { Owner = Application.Current.MainWindow };
                this.mainWindow.Show();
            }
            else
            {
                outputWindowService.OutputLine("Cannot get a DTE service.");
            }
        }
    }
}
