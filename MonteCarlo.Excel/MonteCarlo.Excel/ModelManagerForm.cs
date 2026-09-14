using System;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace MonteCarlo.Excel
{
    public class ModelManagerForm : Form
    {
        private readonly ListView listView;


        public ModelManagerForm()
        {
            Text =
                "Monte Carlo Model Manager";

            Width =
                1150;

            Height =
                540;

            StartPosition =
                FormStartPosition.CenterScreen;

            FormBorderStyle =
                FormBorderStyle.FixedDialog;

            MaximizeBox =
                false;

            MinimizeBox =
                false;


            // =====================================================
            // TITLE
            // =====================================================

            Label lblTitle =
                new Label
                {
                    Text =
                        "Monte Carlo Model",

                    Left =
                        20,

                    Top =
                        15,

                    Width =
                        300,

                    Height =
                        30
                };


            Controls.Add(
                lblTitle);


            // =====================================================
            // LIST VIEW
            // =====================================================

            listView =
                new ListView
                {
                    Left =
                        20,

                    Top =
                        55,

                    Width =
                        1090,

                    Height =
                        370,

                    View =
                        View.Details,

                    FullRowSelect =
                        true,

                    GridLines =
                        true,

                    MultiSelect =
                        false,

                    HideSelection =
                        false
                };


            listView.Columns.Add(
                "Type",
                110);

            listView.Columns.Add(
                "Name",
                180);

            listView.Columns.Add(
                "Sheet",
                120);

            listView.Columns.Add(
                "Cell",
                80);

            listView.Columns.Add(
                "Distribution",
                120);

            listView.Columns.Add(
                "Parameters",
                460);


            Controls.Add(
                listView);


            // =====================================================
            // BUTTONS
            // =====================================================

            Button btnGoToCell =
                new Button
                {
                    Text =
                        "Go to Cell",

                    Left =
                        20,

                    Top =
                        445,

                    Width =
                        110
                };


            Button btnEdit =
                new Button
                {
                    Text =
                        "Edit",

                    Left =
                        145,

                    Top =
                        445,

                    Width =
                        100
                };


            Button btnDelete =
                new Button
                {
                    Text =
                        "Delete",

                    Left =
                        260,

                    Top =
                        445,

                    Width =
                        100
                };


            Button btnRefresh =
                new Button
                {
                    Text =
                        "Refresh",

                    Left =
                        375,

                    Top =
                        445,

                    Width =
                        100
                };


            Button btnClose =
                new Button
                {
                    Text =
                        "Close",

                    Left =
                        1010,

                    Top =
                        445,

                    Width =
                        100
                };


            btnGoToCell.Click +=
                BtnGoToCell_Click;


            btnEdit.Click +=
                BtnEdit_Click;


            btnDelete.Click +=
                BtnDelete_Click;


            btnRefresh.Click +=
                (_, _) =>
                {
                    LoadModelIntoList();
                };


            btnClose.Click +=
                (_, _) =>
                {
                    Close();
                };


            Controls.Add(
                btnGoToCell);

            Controls.Add(
                btnEdit);

            Controls.Add(
                btnDelete);

            Controls.Add(
                btnRefresh);

            Controls.Add(
                btnClose);


            listView.DoubleClick +=
                (_, _) =>
                {
                    GoToSelectedCell();
                };


            LoadModelIntoList();
        }


        // =========================================================
        // LOAD MODEL
        // =========================================================

        private void LoadModelIntoList()
        {
            try
            {
                WorkbookPersistence
                    .LoadModel();


                listView.Items.Clear();


                // -------------------------------------------------
                // ASSUMPTIONS
                // -------------------------------------------------

                foreach (
                    AssumptionDefinition assumption
                    in SimulationModel.Assumptions)
                {
                    ListViewItem item =
                        new ListViewItem(
                            "Assumption");


                    item.SubItems.Add(
                        GetAssumptionName(
                            assumption));


                    item.SubItems.Add(
                        assumption.SheetName);


                    item.SubItems.Add(
                        assumption.CellAddress);


                    item.SubItems.Add(
                        assumption
                            .Distribution
                            .ToString());


                    item.SubItems.Add(
                        GetParameterText(
                            assumption));


                    item.Tag =
                        assumption;


                    listView.Items.Add(
                        item);
                }


                // -------------------------------------------------
                // FORECASTS
                // -------------------------------------------------

                foreach (
                    ForecastDefinition forecast
                    in SimulationModel.Forecasts)
                {
                    ListViewItem item =
                        new ListViewItem(
                            "Forecast");


                    item.SubItems.Add(
                        GetForecastName(
                            forecast));


                    item.SubItems.Add(
                        forecast.SheetName);


                    item.SubItems.Add(
                        forecast.CellAddress);


                    item.SubItems.Add(
                        "-");


                    item.SubItems.Add(
                        "-");


                    item.Tag =
                        forecast;


                    listView.Items.Add(
                        item);
                }
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // ASSUMPTION NAME
        // =========================================================

        private static string GetAssumptionName(
            AssumptionDefinition assumption)
        {
            if (
                !string.IsNullOrWhiteSpace(
                    assumption.Name))
            {
                return
                    assumption.Name;
            }


            return
                $"{assumption.SheetName}!" +
                $"{assumption.CellAddress}";
        }


        // =========================================================
        // FORECAST NAME
        // =========================================================

        private static string GetForecastName(
            ForecastDefinition forecast)
        {
            if (
                !string.IsNullOrWhiteSpace(
                    forecast.Name))
            {
                return
                    forecast.Name;
            }


            return
                $"{forecast.SheetName}!" +
                $"{forecast.CellAddress}";
        }


        // =========================================================
        // PARAMETER DISPLAY
        // =========================================================

        private static string GetParameterText(
            AssumptionDefinition assumption)
        {
            switch (
                assumption.Distribution)
            {
                // -------------------------------------------------
                // NORMAL
                // -------------------------------------------------

                case DistributionType.Normal:

                    return
                        $"Mean={assumption.Parameter1:N2}, " +
                        $"SD={assumption.Parameter2:N2}";


                // -------------------------------------------------
                // TRIANGULAR
                // -------------------------------------------------

                case DistributionType.Triangular:

                    return
                        $"Min={assumption.Parameter1:N2}, " +
                        $"Mode={assumption.Parameter2:N2}, " +
                        $"Max={assumption.Parameter3:N2}";


                // -------------------------------------------------
                // PERT
                // -------------------------------------------------

                case DistributionType.Pert:

                    return
                        $"Min={assumption.Parameter1:N2}, " +
                        $"Most Likely={assumption.Parameter2:N2}, " +
                        $"Max={assumption.Parameter3:N2}";


                // -------------------------------------------------
                // UNIFORM
                // -------------------------------------------------

                case DistributionType.Uniform:

                    return
                        $"Min={assumption.Parameter1:N2}, " +
                        $"Max={assumption.Parameter2:N2}";


                // -------------------------------------------------
                // LOGNORMAL
                // -------------------------------------------------

                case DistributionType.Lognormal:
                    {
                        double logMean =
                            assumption.Parameter1;


                        double logSd =
                            assumption.Parameter2;


                        double median =
                            Math.Exp(
                                logMean);


                        double mean =
                            Math.Exp(
                                logMean +
                                (
                                    logSd *
                                    logSd
                                )
                                /
                                2.0);


                        double variance =
                            (
                                Math.Exp(
                                    logSd *
                                    logSd)
                                -
                                1.0
                            )
                            *
                            Math.Exp(
                                2.0 *
                                logMean
                                +
                                logSd *
                                logSd);


                        double sd =
                            Math.Sqrt(
                                variance);


                        return
                            $"Mean={mean:N2}, " +
                            $"Median={median:N2}, " +
                            $"SD={sd:N2}, " +
                            $"Log μ={logMean:N4}, " +
                            $"σ={logSd:N4}";
                    }


                // -------------------------------------------------
                // BETA
                // -------------------------------------------------

                case DistributionType.Beta:
                    {
                        double minimum =
                            assumption.Parameter1;


                        double maximum =
                            assumption.Parameter2;


                        double alpha =
                            assumption.Parameter3;


                        double beta =
                            assumption.Parameter4;


                        if (
                            alpha <= 0
                            ||
                            beta <= 0
                            ||
                            maximum <= minimum)
                        {
                            return
                                $"Min={minimum:N2}, " +
                                $"Max={maximum:N2}, " +
                                $"Alpha={alpha:N3}, " +
                                $"Beta={beta:N3}";
                        }


                        double mean =
                            minimum
                            +
                            (
                                maximum -
                                minimum
                            )
                            *
                            alpha
                            /
                            (
                                alpha +
                                beta
                            );


                        double variance =
                            Math.Pow(
                                maximum -
                                minimum,
                                2)
                            *
                            (
                                alpha *
                                beta
                            )
                            /
                            (
                                Math.Pow(
                                    alpha +
                                    beta,
                                    2)
                                *
                                (
                                    alpha +
                                    beta +
                                    1.0
                                )
                            );


                        double sd =
                            Math.Sqrt(
                                variance);


                        return
                            $"Min={minimum:N2}, " +
                            $"Max={maximum:N2}, " +
                            $"Alpha={alpha:N3}, " +
                            $"Beta={beta:N3}, " +
                            $"Mean≈{mean:N2}, " +
                            $"SD≈{sd:N2}";
                    }


                default:

                    return
                        "";
            }
        }


        // =========================================================
        // GO TO CELL
        // =========================================================

        private void BtnGoToCell_Click(
            object? sender,
            EventArgs e)
        {
            GoToSelectedCell();
        }


        private void GoToSelectedCell()
        {
            if (listView.SelectedItems.Count == 0)
            {
                MessageBox.Show(
                    "Select an item first.",
                    "Monte Carlo");

                return;
            }


            try
            {
                ListViewItem selectedItem =
                    listView.SelectedItems[0];


                string sheetName =
                    selectedItem
                        .SubItems[2]
                        .Text;


                string cellAddress =
                    selectedItem
                        .SubItems[3]
                        .Text;


                dynamic excelApp =
                    ExcelDnaUtil.Application;


                dynamic sheet =
                    excelApp.Worksheets[
                        sheetName];


                sheet.Activate();


                dynamic cell =
                    sheet.Range[
                        cellAddress];


                cell.Select();
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // EDIT
        // =========================================================

        private void BtnEdit_Click(
            object? sender,
            EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                MessageBox.Show(
                    "Select an item first.",
                    "Monte Carlo");

                return;
            }


            ListViewItem selectedItem =
                listView.SelectedItems[0];


            if (
                selectedItem.Tag
                is AssumptionDefinition assumption)
            {
                EditAssumption(
                    assumption);

                return;
            }


            if (
                selectedItem.Tag
                is ForecastDefinition forecast)
            {
                EditForecast(
                    forecast);

                return;
            }
        }


        // =========================================================
        // EDIT ASSUMPTION
        // =========================================================

        private void EditAssumption(
            AssumptionDefinition assumption)
        {
            try
            {
                using AssumptionForm form =
                    new AssumptionForm(
                        $"{assumption.SheetName}!" +
                        $"{assumption.CellAddress}",
                        assumption);


                if (form.ShowDialog()
                    != DialogResult.OK)
                {
                    return;
                }


                assumption.Name =
                    form.AssumptionName;


                assumption.Distribution =
                    form.Distribution;


                assumption.Parameter1 =
                    form.Parameter1;


                assumption.Parameter2 =
                    form.Parameter2;


                assumption.Parameter3 =
                    form.Parameter3;


                assumption.Parameter4 =
                    form.Parameter4;


                WorkbookPersistence
                    .SaveModel();


                LoadModelIntoList();
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // EDIT FORECAST
        // =========================================================

        private void EditForecast(
            ForecastDefinition forecast)
        {
            try
            {
                string currentName =
                    GetForecastName(
                        forecast);


                string? newName =
                    ShowNameDialog(
                        "Edit Forecast",
                        "Forecast name:",
                        currentName);


                if (newName == null)
                {
                    return;
                }


                newName =
                    newName.Trim();


                if (string.IsNullOrWhiteSpace(
                        newName))
                {
                    newName =
                        $"{forecast.SheetName}!" +
                        $"{forecast.CellAddress}";
                }


                forecast.Name =
                    newName;


                WorkbookPersistence
                    .SaveModel();


                LoadModelIntoList();
            }
            catch (Exception ex)
            {
                ShowError(
                    ex);
            }
        }


        // =========================================================
        // DELETE
        // =========================================================

        private void BtnDelete_Click(
            object? sender,
            EventArgs e)
        {
            if (listView.SelectedItems.Count == 0)
            {
                MessageBox.Show(
                    "Select an item first.",
                    "Monte Carlo");

                return;
            }


            ListViewItem selectedItem =
                listView.SelectedItems[0];


            // -----------------------------------------------------
            // DELETE ASSUMPTION
            // -----------------------------------------------------

            if (
                selectedItem.Tag
                is AssumptionDefinition assumption)
            {
                string assumptionName =
                    GetAssumptionName(
                        assumption);


                DialogResult confirmation =
                    MessageBox.Show(
                        $"Delete assumption '{assumptionName}'?\n\n" +
                        $"Cell: {assumption.SheetName}!" +
                        $"{assumption.CellAddress}",
                        "Monte Carlo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);


                if (confirmation !=
                    DialogResult.Yes)
                {
                    return;
                }


                SimulationModel.Assumptions
                    .RemoveAll(
                        x =>
                            string.Equals(
                                x.SheetName,
                                assumption.SheetName,
                                StringComparison.OrdinalIgnoreCase)
                            &&
                            string.Equals(
                                x.CellAddress,
                                assumption.CellAddress,
                                StringComparison.OrdinalIgnoreCase));


                WorkbookPersistence
                    .SaveModel();


                LoadModelIntoList();


                return;
            }


            // -----------------------------------------------------
            // DELETE FORECAST
            // -----------------------------------------------------

            if (
                selectedItem.Tag
                is ForecastDefinition forecast)
            {
                string forecastName =
                    GetForecastName(
                        forecast);


                DialogResult confirmation =
                    MessageBox.Show(
                        $"Delete forecast '{forecastName}'?\n\n" +
                        $"Cell: {forecast.SheetName}!" +
                        $"{forecast.CellAddress}",
                        "Monte Carlo",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);


                if (confirmation !=
                    DialogResult.Yes)
                {
                    return;
                }


                SimulationModel.Forecasts
                    .RemoveAll(
                        x =>
                            string.Equals(
                                x.SheetName,
                                forecast.SheetName,
                                StringComparison.OrdinalIgnoreCase)
                            &&
                            string.Equals(
                                x.CellAddress,
                                forecast.CellAddress,
                                StringComparison.OrdinalIgnoreCase));


                WorkbookPersistence
                    .SaveModel();


                LoadModelIntoList();
            }
        }


        // =========================================================
        // NAME DIALOG
        // =========================================================

        private static string? ShowNameDialog(
            string title,
            string labelText,
            string currentValue)
        {
            using Form dialog =
                new Form
                {
                    Text =
                        title,

                    Width =
                        400,

                    Height =
                        180,

                    StartPosition =
                        FormStartPosition.CenterParent,

                    FormBorderStyle =
                        FormBorderStyle.FixedDialog,

                    MaximizeBox =
                        false,

                    MinimizeBox =
                        false
                };


            Label label =
                new Label
                {
                    Text =
                        labelText,

                    Left =
                        20,

                    Top =
                        25,

                    Width =
                        120
                };


            TextBox textBox =
                new TextBox
                {
                    Left =
                        140,

                    Top =
                        20,

                    Width =
                        210,

                    Text =
                        currentValue
                };


            Button btnOK =
                new Button
                {
                    Text =
                        "OK",

                    Left =
                        190,

                    Top =
                        70,

                    Width =
                        75,

                    DialogResult =
                        DialogResult.OK
                };


            Button btnCancel =
                new Button
                {
                    Text =
                        "Cancel",

                    Left =
                        275,

                    Top =
                        70,

                    Width =
                        75,

                    DialogResult =
                        DialogResult.Cancel
                };


            dialog.Controls.Add(
                label);

            dialog.Controls.Add(
                textBox);

            dialog.Controls.Add(
                btnOK);

            dialog.Controls.Add(
                btnCancel);


            dialog.AcceptButton =
                btnOK;


            dialog.CancelButton =
                btnCancel;


            if (dialog.ShowDialog()
                != DialogResult.OK)
            {
                return null;
            }


            return
                textBox.Text;
        }


        // =========================================================
        // ERROR
        // =========================================================

        private static void ShowError(
            Exception ex)
        {
            MessageBox.Show(
                ex.ToString(),
                "Monte Carlo Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}