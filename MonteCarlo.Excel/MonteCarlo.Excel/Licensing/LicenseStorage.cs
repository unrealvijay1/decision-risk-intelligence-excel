using System;
using System.IO;

namespace MonteCarlo.Excel.Licensing
{
    public static class LicenseStorage
    {
        private const string LicenseFileName =
            "license.dat";


        // =========================================================
        // SAVE LICENSE KEY
        // =========================================================

        public static void SaveLicenseKey(
            string licenseKey)
        {
            if (string.IsNullOrWhiteSpace(
                    licenseKey))
            {
                throw new ArgumentException(
                    "License key cannot be empty.",
                    nameof(licenseKey));
            }


            string path =
                GetLicenseFilePath();


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
                licenseKey.Trim());
        }


        // =========================================================
        // LOAD LICENSE KEY
        // =========================================================

        public static string? LoadLicenseKey()
        {
            string path =
                GetLicenseFilePath();


            if (!File.Exists(
                    path))
            {
                return null;
            }


            try
            {
                string value =
                    File.ReadAllText(
                        path);


                if (string.IsNullOrWhiteSpace(
                        value))
                {
                    return null;
                }


                return
                    value.Trim();
            }
            catch
            {
                return null;
            }
        }


        // =========================================================
        // HAS SAVED LICENSE
        // =========================================================

        public static bool HasSavedLicense()
        {
            string? key =
                LoadLicenseKey();


            return
                !string.IsNullOrWhiteSpace(
                    key);
        }


        // =========================================================
        // CLEAR LICENSE
        // =========================================================

        public static void ClearLicense()
        {
            string path =
                GetLicenseFilePath();


            if (!File.Exists(
                    path))
            {
                return;
            }


            try
            {
                File.Delete(
                    path);
            }
            catch
            {
                // Ignore delete errors for now.
            }
        }


        // =========================================================
        // LICENSE FILE LOCATION
        // =========================================================

        private static string GetLicenseFilePath()
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
                    LicenseFileName);
        }
    }
}