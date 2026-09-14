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
        // DEVELOPMENT TEST KEYS
        //
        // IMPORTANT:
        // These are only temporary local-development keys.
        // They must NOT be used in the commercial release.
        // =========================================================

        private const string DevelopmentProfessionalKey =
            "MC-PRO-DEV-2026";


        private const string DevelopmentEnterpriseKey =
            "MC-ENT-DEV-2026";


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
            // CHECK SAVED ACTIVATED LICENSE
            // -----------------------------------------------------

            string? savedLicenseKey =
                LicenseStorage.LoadLicenseKey();


            if (!string.IsNullOrWhiteSpace(
                    savedLicenseKey))
            {
                LicenseInfo? activatedLicense =
                    ValidateDevelopmentLicenseKey(
                        savedLicenseKey);


                if (
                    activatedLicense != null
                    &&
                    activatedLicense.IsValid)
                {
                    return
                        activatedLicense;
                }
            }


            // -----------------------------------------------------
            // FALL BACK TO TRIAL
            // -----------------------------------------------------

            return
                GetTrialLicense();
        }


        // =========================================================
        // ACTIVATE
        // =========================================================

        public static LicenseInfo? Activate(
            string licenseKey)
        {
            if (string.IsNullOrWhiteSpace(
                    licenseKey))
            {
                return
                    null;
            }


            string normalizedKey =
                NormalizeLicenseKey(
                    licenseKey);


            // -----------------------------------------------------
            // TEMPORARY DEVELOPMENT VALIDATION
            // -----------------------------------------------------

            LicenseInfo? license =
                ValidateDevelopmentLicenseKey(
                    normalizedKey);


            if (
                license == null
                ||
                !license.IsValid)
            {
                return
                    null;
            }


            // -----------------------------------------------------
            // SAVE ACTIVATED KEY
            // -----------------------------------------------------

            LicenseStorage.SaveLicenseKey(
                normalizedKey);


            return
                license;
        }


        // =========================================================
        // VALIDATE SAVED / ENTERED KEY
        //
        // Later this method will be replaced by an HTTPS call to
        // the production licensing service.
        // =========================================================

        private static LicenseInfo?
            ValidateDevelopmentLicenseKey(
                string licenseKey)
        {
            string normalizedKey =
                NormalizeLicenseKey(
                    licenseKey);


            // -----------------------------------------------------
            // PROFESSIONAL TEST LICENSE
            // -----------------------------------------------------

            if (
                string.Equals(
                    normalizedKey,
                    DevelopmentProfessionalKey,
                    StringComparison.OrdinalIgnoreCase))
            {
                return
                    new LicenseInfo
                    {
                        Type =
                            LicenseType.Professional,

                        Status =
                            LicenseStatus.Active,

                        LicensedTo =
                            "Development Professional User",

                        ExpiryDate =
                            DateTime.Today
                                .AddYears(
                                    1)
                    };
            }


            // -----------------------------------------------------
            // ENTERPRISE TEST LICENSE
            // -----------------------------------------------------

            if (
                string.Equals(
                    normalizedKey,
                    DevelopmentEnterpriseKey,
                    StringComparison.OrdinalIgnoreCase))
            {
                return
                    new LicenseInfo
                    {
                        Type =
                            LicenseType.Enterprise,

                        Status =
                            LicenseStatus.Active,

                        LicensedTo =
                            "Development Enterprise User",

                        ExpiryDate =
                            DateTime.Today
                                .AddYears(
                                    1)
                    };
            }


            return
                null;
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


            throw new InvalidOperationException(
                "Your Monte Carlo license is not active.");
        }


        // =========================================================
        // DEACTIVATE
        //
        // For now this only clears the locally stored key.
        // Production deactivation will also notify the license
        // server and release the device activation.
        // =========================================================

        public static void Deactivate()
        {
            LicenseStorage.ClearLicense();
        }


        // =========================================================
        // NORMALIZE KEY
        // =========================================================

        private static string NormalizeLicenseKey(
            string licenseKey)
        {
            return
                licenseKey
                    .Trim()
                    .ToUpperInvariant();
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