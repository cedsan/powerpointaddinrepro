using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace UserControls
{
    public partial class UserControl1 : UserControl
    {
        private readonly string[] availableColorThemes = { "Test Dark", "Test Light" };
        
        private string currentTheme;
        private Microsoft.Office.Interop.PowerPoint.Application application;

        public UserControl1(Microsoft.Office.Interop.PowerPoint.Application application)
        {
            InitializeComponent();
            this.application = application;
            currentTheme = "Test Light";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var presentation = application.ActivePresentation;
            currentTheme = availableColorThemes.Where(c => c != currentTheme).First();
            string fileName = Path.Combine(Path.GetTempPath(), "PowerPointAddIn", $"{currentTheme}.xml");

            for (var i = 1; i <= presentation.Slides.Count; i++)
            {
                presentation.Slides[i].ThemeColorScheme.Load(fileName);
            }
        }
    }
}
