using System;

namespace MonteCarlo.Excel.Licensing
{
    public enum LicenseType
    {
        Trial,
        Professional,
        Enterprise,
        Development
    }


    public enum LicenseStatus
    {
        Active,
        Expired,
        Invalid
    }


    public class LicenseInfo
    {
        public LicenseType Type { get; set; }

        public LicenseStatus Status { get; set; }

        public string LicensedTo { get; set; } = "";

        public DateTime? ExpiryDate { get; set; }

        public bool IsValid =>
            Status == LicenseStatus.Active;


        public int DaysRemaining
        {
            get
            {
                if (!ExpiryDate.HasValue)
                {
                    return int.MaxValue;
                }


                double days =
                    (
                        ExpiryDate.Value.Date -
                        DateTime.Today
                    ).TotalDays;


                return Math.Max(
                    0,
                    (int)Math.Ceiling(days));
            }
        }
    }
}