using System.IO;
using System.ComponentModel;
using System.Configuration.Install;
using CompareXMLDocs;
using System.Diagnostics;
using System.Windows.Forms;

namespace InstallerActions
{
    [RunInstaller(true)]
    public partial class Installer1 : System.Configuration.Install.Installer
    {
        public Installer1() : base()
        {
            InitializeComponent();
        }
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);
        }

        public override void Commit(System.Collections.IDictionary savedState)
        {

            base.Commit(savedState);
            _ = PostBuild.UpdateConfigurationFile(Path.GetTempPath() + "BCE\\ZibExtraction\\UpdateConfig.xml", out bool succes, out string logfile);
            this.Context.Parameters.Add("Outcome", succes.ToString());
            if (!succes) MessageBox.Show(new Form { TopMost = true }, "Update configuratie mislukt.\r\n"
                + "Voor details zie logfile " + logfile);
        }
    }
}
