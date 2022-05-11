using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BrfAlbert.Aptus.UI
{
    public partial class AptusBillingUI : Form
    {
        public AptusBillingUI()
        {
            InitializeComponent();
            InitializePeriods();
        }

        private void InitializePeriods()
        {
            periodComboBox.Items.Clear();
            var date = DateTime.UtcNow.AddMonths(-1);
            for (int i = 0; i < 10; i++)
            {
                var period = $"{date:yyyy-MM}";
                periodComboBox.Items.Add(period);
                date = date.AddMonths(-1);
            }
            periodComboBox.SelectedIndex = 0;
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
