using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;

namespace MonteCarlo.Excel.Licensing
{
    public static class TrialService
    {
        // =========================================================
        // SETTINGS
        // =========================================================

        private const int TrialDays =
            30;


        private const string TrialFileName =
            "trial.dat";


        private const string RegistryPath =
            @"Software\MonteCarloExcel";


        private const string RegistryValueName =
            "TrialState";


        // Extra entropy used with Windows DPAPI.
        //
        // This is NOT a secret key.
        // It simply makes the protected payload specific
        // to this application.
        private static readonly byte[] AdditionalEntropy =
            Encoding.UTF8.GetBytes(
                "MonteCarloExcel.TrialState.v1");


        // =========================================================
        // GET TRIAL LICENSE
        // =========================================================

        public static LicenseInfo GetTrialLicense()
        {
            TrialStateResult stateResult =
                LoadOrCreateState();


            // -----------------------------------------------------
            // CORRUPTION / TAMPERING
            // -----------------------------------------------------

            if (!stateResult.IsValid)
            {
                return
                    new LicenseInfo
                    {
                        Type =
                            LicenseType.Trial,

                        Status =
                            LicenseStatus.Expired,

                        LicensedTo =
                            "Trial User",

                        ExpiryDate =
                            stateResult.State.StartDate
                                .AddDays(
                                    TrialDays)
                    };
            }


            TrialState state =
                stateResult.State;


            DateTime today =
                DateTime.Today;


            // =====================================================
            // CLOCK ROLLBACK DETECTION
            // =====================================================

            if (today < state.LastRunDate.Date)
            {
                MarkTrialInvalid(
                    state);


                return
                    new LicenseInfo
                    {
                        Type =
                            LicenseType.Trial,

                        Status =
                            LicenseStatus.Expired,

                        LicensedTo =
                            "Trial User",

                        ExpiryDate =
                            state.StartDate
                                .AddDays(
                                    TrialDays)
                    };
            }


            // =====================================================
            // INVALID FUTURE START DATE
            // =====================================================

            if (state.StartDate.Date > today)
            {
                MarkTrialInvalid(
                    state);


                return
                    new LicenseInfo
                    {
                        Type =
                            LicenseType.Trial,

                        Status =
                            LicenseStatus.Expired,

                        LicensedTo =
                            "Trial User",

                        ExpiryDate =
                            state.StartDate
                                .AddDays(
                                    TrialDays)
                    };
            }


            // =====================================================
            // UPDATE LAST RUN DATE
            // =====================================================

            if (today > state.LastRunDate.Date)
            {
                state.LastRunDate =
                    today;


                SaveState(
                    state);
            }


            // =====================================================
            // EXPIRY
            // =====================================================

            DateTime expiryDate =
                state.StartDate.Date
                    .AddDays(
                        TrialDays);


            LicenseStatus status =
                today <= expiryDate
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
        // LOAD OR CREATE STATE
        // =========================================================

        private static TrialStateResult LoadOrCreateState()
        {
            TrialState? fileState =
                LoadFromFile();


            TrialState? registryState =
                LoadFromRegistry();


            // =====================================================
            // BOTH MISSING
            //
            // First legitimate use.
            // =====================================================

            if (
                fileState == null
                &&
                registryState == null)
            {
                TrialState newState =
                    CreateNewState();


                SaveState(
                    newState);


                return
                    new TrialStateResult
                    {
                        IsValid =
                            true,

                        State =
                            newState
                    };
            }


            // =====================================================
            // FILE MISSING, REGISTRY EXISTS
            //
            // Restore deleted/missing file.
            // =====================================================

            if (
                fileState == null
                &&
                registryState != null)
            {
                SaveToFile(
                    registryState);


                return
                    ValidateState(
                        registryState);
            }


            // =====================================================
            // REGISTRY MISSING, FILE EXISTS
            //
            // Restore deleted/missing registry entry.
            // =====================================================

            if (
                registryState == null
                &&
                fileState != null)
            {
                SaveToRegistry(
                    fileState);


                return
                    ValidateState(
                        fileState);
            }


            // At this point both exist.
            if (
                fileState == null
                ||
                registryState == null)
            {
                return
                    CreateInvalidResult();
            }


            // =====================================================
            // CONSERVATIVE MERGE
            //
            // If the two copies differ, use:
            //
            // Earliest start date
            // Latest last-run date
            //
            // This prevents replacing one copy with an older state
            // to extend the trial.
            // =====================================================

            TrialState mergedState =
                new TrialState
                {
                    TrialId =
                        string.IsNullOrWhiteSpace(
                            fileState.TrialId)
                            ? registryState.TrialId
                            : fileState.TrialId,

                    StartDate =
                        fileState.StartDate <=
                        registryState.StartDate
                            ? fileState.StartDate
                            : registryState.StartDate,

                    LastRunDate =
                        fileState.LastRunDate >=
                        registryState.LastRunDate
                            ? fileState.LastRunDate
                            : registryState.LastRunDate,

                    IsInvalid =
                        fileState.IsInvalid
                        ||
                        registryState.IsInvalid
                };


            // -----------------------------------------------------
            // TRIAL ID MISMATCH
            //
            // Two unrelated trial states should never normally
            // exist for the same Windows user.
            // -----------------------------------------------------

            if (
                !string.IsNullOrWhiteSpace(
                    fileState.TrialId)
                &&
                !string.IsNullOrWhiteSpace(
                    registryState.TrialId)
                &&
                !string.Equals(
                    fileState.TrialId,
                    registryState.TrialId,
                    StringComparison.Ordinal))
            {
                mergedState.IsInvalid =
                    true;
            }


            SaveState(
                mergedState);


            return
                ValidateState(
                    mergedState);
        }


        // =========================================================
        // VALIDATE STATE
        // =========================================================

        private static TrialStateResult ValidateState(
            TrialState state)
        {
            bool valid =
                !state.IsInvalid
                &&
                !string.IsNullOrWhiteSpace(
                    state.TrialId)
                &&
                state.StartDate !=
                    DateTime.MinValue
                &&
                state.LastRunDate !=
                    DateTime.MinValue
                &&
                state.LastRunDate >=
                    state.StartDate;


            return
                new TrialStateResult
                {
                    IsValid =
                        valid,

                    State =
                        state
                };
        }


        // =========================================================
        // NEW TRIAL STATE
        // =========================================================

        private static TrialState CreateNewState()
        {
            DateTime today =
                DateTime.Today;


            return
                new TrialState
                {
                    TrialId =
                        Guid.NewGuid()
                            .ToString(
                                "N"),

                    StartDate =
                        today,

                    LastRunDate =
                        today,

                    IsInvalid =
                        false
                };
        }


        // =========================================================
        // MARK INVALID
        // =========================================================

        private static void MarkTrialInvalid(
            TrialState state)
        {
            state.IsInvalid =
                true;


            SaveState(
                state);
        }


        // =========================================================
        // SAVE BOTH COPIES
        // =========================================================

        private static void SaveState(
            TrialState state)
        {
            SaveToFile(
                state);


            SaveToRegistry(
                state);
        }


        // =========================================================
        // SAVE TO FILE
        // =========================================================

        private static void SaveToFile(
            TrialState state)
        {
            try
            {
                string path =
                    GetTrialFilePath();


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


                string protectedValue =
                    ProtectState(
                        state);


                File.WriteAllText(
                    path,
                    protectedValue);
            }
            catch
            {
                // Registry copy remains available.
            }
        }


        // =========================================================
        // LOAD FROM FILE
        // =========================================================

        private static TrialState? LoadFromFile()
        {
            try
            {
                string path =
                    GetTrialFilePath();


                if (!File.Exists(
                        path))
                {
                    return null;
                }


                string protectedValue =
                    File.ReadAllText(
                        path);


                return
                    UnprotectState(
                        protectedValue);
            }
            catch
            {
                return null;
            }
        }


        // =========================================================
        // SAVE TO REGISTRY
        // =========================================================

        private static void SaveToRegistry(
            TrialState state)
        {
            try
            {
                using RegistryKey? key =
                    Registry.CurrentUser
                        .CreateSubKey(
                            RegistryPath);


                if (key == null)
                {
                    return;
                }


                string protectedValue =
                    ProtectState(
                        state);


                key.SetValue(
                    RegistryValueName,
                    protectedValue,
                    RegistryValueKind.String);
            }
            catch
            {
                // File copy remains available.
            }
        }


        // =========================================================
        // LOAD FROM REGISTRY
        // =========================================================

        private static TrialState? LoadFromRegistry()
        {
            try
            {
                using RegistryKey? key =
                    Registry.CurrentUser
                        .OpenSubKey(
                            RegistryPath);


                if (key == null)
                {
                    return null;
                }


                string? protectedValue =
                    key.GetValue(
                        RegistryValueName)
                    as string;


                if (string.IsNullOrWhiteSpace(
                        protectedValue))
                {
                    return null;
                }


                return
                    UnprotectState(
                        protectedValue);
            }
            catch
            {
                return null;
            }
        }


        // =========================================================
        // PROTECT STATE WITH WINDOWS DPAPI
        // =========================================================

        private static string ProtectState(
            TrialState state)
        {
            string json =
                JsonSerializer.Serialize(
                    state);


            byte[] plainBytes =
                Encoding.UTF8.GetBytes(
                    json);


            byte[] encryptedBytes =
                ProtectedData.Protect(
                    plainBytes,
                    AdditionalEntropy,
                    DataProtectionScope.CurrentUser);


            return
                Convert.ToBase64String(
                    encryptedBytes);
        }


        // =========================================================
        // UNPROTECT STATE
        // =========================================================

        private static TrialState? UnprotectState(
            string protectedValue)
        {
            if (string.IsNullOrWhiteSpace(
                    protectedValue))
            {
                return null;
            }


            try
            {
                byte[] encryptedBytes =
                    Convert.FromBase64String(
                        protectedValue);


                byte[] plainBytes =
                    ProtectedData.Unprotect(
                        encryptedBytes,
                        AdditionalEntropy,
                        DataProtectionScope.CurrentUser);


                string json =
                    Encoding.UTF8.GetString(
                        plainBytes);


                return
                    JsonSerializer.Deserialize<
                        TrialState>(
                            json);
            }
            catch
            {
                return null;
            }
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


        // =========================================================
        // INVALID RESULT
        // =========================================================

        private static TrialStateResult CreateInvalidResult()
        {
            TrialState state =
                CreateNewState();


            state.IsInvalid =
                true;


            return
                new TrialStateResult
                {
                    IsValid =
                        false,

                    State =
                        state
                };
        }


        // =========================================================
        // INTERNAL STATE
        // =========================================================

        private sealed class TrialState
        {
            public string TrialId { get; set; } =
                "";


            public DateTime StartDate { get; set; }


            public DateTime LastRunDate { get; set; }


            public bool IsInvalid { get; set; }
        }


        // =========================================================
        // INTERNAL RESULT
        // =========================================================

        private sealed class TrialStateResult
        {
            public bool IsValid { get; set; }


            public TrialState State { get; set; } =
                new TrialState();
        }
    }
}