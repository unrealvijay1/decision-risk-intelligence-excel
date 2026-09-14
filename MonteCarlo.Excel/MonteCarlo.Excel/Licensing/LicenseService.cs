using System;

namespace MonteCarlo.Excel.Licensing
{
    public static class LicenseService
    {
        // =========================================================
        // DEVELOPMENT SETTINGS
        // =========================================================

        // TRUE:
        //   Development machine has unrestricted access.
        //
        // FALSE:
        //   Normal commercial licensing flow:
        //   Signed license -> Trial fallback.
        //
        // Keep FALSE when testing the customer experience.
        private const bool DevelopmentMode =
            false;


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
            // CHECK SAVED SIGNED CUSTOMER LICENSE
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
                    // Return even if expired.
                    //
                    // This allows the License screen to show:
                    //
                    // Professional / Enterprise
                    // Expired
                    // Customer name
                    // Expiry date

                    return
                        savedLicense;
                }


                // -------------------------------------------------
                // INVALID / TAMPERED SAVED LICENSE
                // -------------------------------------------------

                // If the stored license no longer passes RSA
                // validation, remove it.
                LicenseStorage.ClearLicense();
            }


            // -----------------------------------------------------
            // NO CUSTOMER LICENSE
            //
            // Use protected local trial.
            // -----------------------------------------------------

            return
                TrialService.GetTrialLicense();
        }


        // =========================================================
        // ACTIVATE LICENSE
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


            // -----------------------------------------------------
            // IMPORTANT
            //
            // Signed license codes are CASE-SENSITIVE.
            //
            // Do not uppercase or otherwise modify the code.
            // -----------------------------------------------------

            string normalizedCode =
                NormalizeLicenseCode(
                    licenseCode);


            // -----------------------------------------------------
            // VERIFY SIGNATURE AND PAYLOAD
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
            // DO NOT STORE EXPIRED LICENSE
            // -----------------------------------------------------

            if (!license.IsValid)
            {
                return
                    license;
            }


            // -----------------------------------------------------
            // SAVE VALID SIGNED LICENSE
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
        // ENSURE ACCESS
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
                if (
                    license.Type ==
                    LicenseType.Trial)
                {
                    throw new InvalidOperationException(
                        "Your 30-day Monte Carlo trial has expired.");
                }


                throw new InvalidOperationException(
                    "Your Monte Carlo license has expired.");
            }


            throw new InvalidOperationException(
                "Your Monte Carlo license is not active.");
        }


        // =========================================================
        // DEACTIVATE LICENSE
        // =========================================================

        public static void Deactivate()
        {
            // Removes the signed customer license.
            //
            // IMPORTANT:
            // This does NOT reset the trial.
            //
            // If the original trial has already expired,
            // deactivation returns the user to that expired trial.

            LicenseStorage.ClearLicense();
        }


        // =========================================================
        // NORMALIZE LICENSE CODE
        // =========================================================

        private static string NormalizeLicenseCode(
            string licenseCode)
        {
            // Only remove whitespace before/after the entire code.
            //
            // Do NOT call ToUpperInvariant().
            // Base64URL data inside the signed license is
            // case-sensitive.

            return
                licenseCode.Trim();
        }
    }
}