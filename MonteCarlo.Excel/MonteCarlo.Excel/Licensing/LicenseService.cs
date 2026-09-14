using System;
using System.IO;

namespace MonteCarlo.Excel.Licensing
{
    public static class LicenseService
    {
        // =========================================================
        // DEVELOPMENT SETTINGS
        // =========================================================

        private const bool DevelopmentMode =
            false;


        private const int TrialDays =
            30;


        private const string TrialFileName =
            "trial.dat";


        // =========================================================
        // GET CURRENT LICENSE
        // =========================================================

        public static LicenseInfo GetCurrentLicense()
        {
            // -----------------------------------------------------
            // DEVELOPMENT OVERRIDE
            // -----------------------------------------------------

            if (DevelopmentMode)
            {
                return
                    new LicenseInfo
                    {
                        Type =
                            LicenseType.Development,

                        Status =
                            LicenseStatus.Active,

                        LicensedTo =
                            "Development",

                        ExpiryDate =
                            null
                    };
            }


            // -----------------------------------------------------
            // CHECK SAVED SIGNED LICENSE
            // -----------------------------------------------------

            string? savedLicenseCode =
                LicenseStorage.LoadLicenseKey();


            if (!string.IsNullOrWhiteSpace(
                    savedLicenseCode))
            {
                LicenseInfo? savedLicense =
                    LocalLicenseValidator.Validate(
                        savedLicenseCode);


                if (savedLicense != null)
                {
                    // Important:
                    // Return the actual license even when expired.
                    //
                    // This allows LicenseForm to show:
                    // Professional / Enterprise
                    // Expired
                    // Customer name
                    // Expiry date

                    return
                        savedLicense;
                }


                // Invalid or tampered stored license.
                //
                // Remove it so future checks fall back to the trial.
                LicenseStorage.ClearLicense();
            }


            // -----------------------------------------------------
            // FALL BACK TO TRIAL
            // -----------------------------------------------------

            return
                GetTrialLicense();
        }


        // =========================================================
        // ACTIVATE SIGNED LICENSE
        // =========================================================

        public static LicenseInfo? Activate(
            string licenseCode)
        {
            if (string.IsNullOrWhiteSpace(
                    licenseCode))
            {
                return
                    null;
            }


            string normalizedCode =
                NormalizeLicenseCode(
                    licenseCode);


            // -----------------------------------------------------
            // VERIFY SIGNATURE + PAYLOAD
            // -----------------------------------------------------

            LicenseInfo? license =
                LocalLicenseValidator.Validate(
                    normalizedCode);


            if (license == null)
            {
                return
                    null;
            }


            // -----------------------------------------------------
            // DO NOT ACTIVATE AN EXPIRED LICENSE
            // -----------------------------------------------------

            if (!license.IsValid)
            {
                return
                    license;
            }


            // -----------------------------------------------------
            // SAVE SIGNED LICENSE LOCALLY
            // -----------------------------------------------------

            LicenseStorage.SaveLicenseKey(
                normalizedCode);


            return
                license;
        }


        // =========================================================
        // HAS ACCESS
        // =========================================================

        public static bool HasAccess()
        {
            LicenseInfo license =
                GetCurrentLicense();


            return
                license.IsValid;
        }


        // =========================================================
        // REQUIRE ACCESS
        // =========================================================

        public static void EnsureAccess()
        {
            LicenseInfo license =
                GetCurrentLicense();


            if (license.IsValid)
            {
                return;
            }


            if (
                license.Status ==
                LicenseStatus.Expired)
            {
                throw new InvalidOperationException(
                    "Your Monte Carlo license or trial has expired.");
            }


            throw new InvalidOperationException(
                "Your Monte Carlo license is not active.");
        }


        // =========================================================
        // DEACTIVATE
        // =========================================================

        public static void Deactivate()
        {
            LicenseStorage.ClearLicense();
        }


        // =========================================================
        // NORMALIZE LICENSE CODE
        //
        // IMPORTANT:
        //
        // Signed license codes are case-sensitive because Base64URL
        // payload/signature data is case-sensitive.
        //
        // Therefore DO NOT call ToUpperInvariant().
        // =========================================================

        private static string NormalizeLicenseCode(
            string licenseCode)
        {
            return
                licenseCode
                    .Trim();
        }


        // =========================================================
        // TRIAL LICENSE
        // =========================================================

        private static LicenseInfo GetTrialLicense()
        {
            DateTime trialStart =
                GetOrCreateTrialStartDate();


            DateTime expiryDate =
                trialStart
                    .AddDays(
                        TrialDays);


            LicenseStatus status =
                DateTime.Today <= expiryDate
                    ? LicenseStatus.Active
                    : LicenseStatus.Expired;


            return
                new LicenseInfo
                {
                    Type =
                        LicenseType.Trial,

                    Status =
                        status,

                    LicensedTo =
                        "Trial User",

                    ExpiryDate =
                        expiryDate
                };
        }


        // =========================================================
        // TRIAL START DATE
        // =========================================================

        private static DateTime GetOrCreateTrialStartDate()
        {
            string path =
                GetTrialFilePath();


            if (File.Exists(
                    path))
            {
                try
                {
                    string storedValue =
                        File.ReadAllText(
                            path);


                    if (
                        DateTime.TryParse(
                            storedValue,
                            out DateTime trialStart))
                    {
                        return
                            trialStart.Date;
                    }
                }
                catch
                {
                    // Development-stage fallback.
                }
            }


            DateTime newTrialStart =
                DateTime.Today;


            string? directory =
                Path.GetDirectoryName(
                    path);


            if (
                !string.IsNullOrWhiteSpace(
                    directory))
            {
                Directory.CreateDirectory(
                    directory);
            }


            File.WriteAllText(
                path,
                newTrialStart.ToString(
                    "O"));


            return
                newTrialStart;
        }


        // =========================================================
        // TRIAL FILE LOCATION
        // =========================================================

        private static string GetTrialFilePath()
        {
            string appData =
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData);


            string folder =
                Path.Combine(
                    appData,
                    "MonteCarloExcel");


            return
                Path.Combine(
                    folder,
                    TrialFileName);
        }
    }
}