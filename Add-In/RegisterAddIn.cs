using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Xml;
using System.EnterpriseServices.Internal;
using Microsoft.MediaCenter.Hosting;
using VmcController.AddIn.Properties;
using System.IO;
using VmcController.AddIn.Commands;

namespace VmcController.AddIn
{
    [RunInstaller(true)]
    public partial class RegisterAddIn : Installer
    {        

        /// <summary>
        /// Initializes a new instance of the <see cref="RegisterAddIn"/> class.
        /// </summary>
        public RegisterAddIn()
        {
            InitializeComponent();
        }

        /// <summary>
        /// When overridden in a derived class, performs the installation.
        /// </summary>
        /// <param name="stateSaver">An <see cref="T:System.Collections.IDictionary"/> used to save information needed to perform a commit, rollback, or uninstall operation.</param>
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);

            //  Try removing the application definition first, ignore any errors
            try
            {
                ApplicationContext.RegisterApplication(
                    new XmlTextReader(new System.IO.StringReader(VmcController.AddIn.Properties.Resources.Registration)),
                    true, true, string.Empty);
            }
            catch { }

            //  Register the appliation definition
            try
            {
                ApplicationContext.RegisterApplication(
                    new XmlTextReader(new System.IO.StringReader(VmcController.AddIn.Properties.Resources.Registration)),
                    false, true, string.Empty);
            }
            catch (Microsoft.MediaCenter.ApplicationAlreadyRegisteredException)
            {
                //  Already registered
            }
            catch (Exception ex)
            {
                throw new InstallException("Error registering MCE application", ex);
            }
        }

        /// <summary>
        /// When overridden in a derived class, restores the pre-installation state of the computer.
        /// </summary>
        /// <param name="savedState">An <see cref="T:System.Collections.IDictionary"/> that contains the pre-installation state of the computer.</param>
        public override void Rollback(IDictionary savedState)
        {
            try
            {
                ApplicationContext.RegisterApplication(
                    new XmlTextReader(new System.IO.StringReader(Resources.Registration)),
                    true, true, string.Empty);
            }
            catch { }
            finally
            {
                base.Rollback(savedState);
            }
        }

        /// <summary>
        /// When overridden in a derived class, removes an installation.
        /// </summary>
        /// <param name="savedState">An <see cref="T:System.Collections.IDictionary"/> that contains the state of the computer after the installation was complete.</param>
        public override void Uninstall(IDictionary savedState)
        {
            try
            {
                ApplicationContext.RegisterApplication(
                    new XmlTextReader(new System.IO.StringReader(Resources.Registration)),
                    true, true, string.Empty);
            }
            catch { }
            finally
            {
                base.Uninstall(savedState);
            }
        }
    }
}
