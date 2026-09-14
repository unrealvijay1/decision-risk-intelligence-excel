using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MonteCarlo.Excel.Licensing
{
    public static class LocalLicenseValidator
    {
        private const string LicensePrefix =
            "MC1";


        private const string PublicKeyResourceName =
            "MonteCarlo.Excel.Licensing.public-key.pem";


        // =========================================================
        // VALIDATE
        // =========================================================

        public static LicenseInfo? Validate(
            string licenseCode)
        {
            if (string.IsNullOrWhiteSpace(
                    licenseCode))
            {
                return null;
            }


            try
            {
                string[] parts =
                    licenseCode
                        .Trim()
                        .Split('.');


                if (parts.Length != 3)
                {
                    return null;
                }


                if (!string.Equals(
                        parts[0],
                        LicensePrefix,
                        StringComparison.Ordinal))
                {
                    return null;
                }


                byte[] payloadBytes =
                    Base64UrlDecode(
                        parts[1]);


                byte[] signatureBytes =
                    Base64UrlDecode(
                        parts[2]);


                // =================================================
                // VERIFY SIGNATURE
                // =================================================

                string publicKeyPem =
                    LoadPublicKey();


                using RSA rsa =
                    RSA.Create();


                rsa.ImportFromPem(
                    publicKeyPem);


                bool signatureValid =
                    rsa.VerifyData(
                        payloadBytes,
                        signatureBytes,
                        HashAlgorithmName.SHA256,
                        RSASignaturePadding.Pkcs1);


                if (!signatureValid)
                {
                    return null;
                }


                // =================================================
                // READ PAYLOAD
                // =================================================

                string payloadJson =
                    Encoding.UTF8.GetString(
                        payloadBytes);


                LicensePayload? payload =
                    JsonSerializer.Deserialize<
                        LicensePayload>(
                            payloadJson);


                if (payload == null)
                {
                    return null;
                }


                if (string.IsNullOrWhiteSpace(
                        payload.LicenseId))
                {
                    return null;
                }


                if (string.IsNullOrWhiteSpace(
                        payload.Customer))
                {
                    return null;
                }


                // =================================================
                // PARSE EDITION
                // =================================================

                LicenseType licenseType;


                if (
                    string.Equals(
                        payload.Edition,
                        "Professional",
                        StringComparison.OrdinalIgnoreCase))
                {
                    licenseType =
                        LicenseType.Professional;
                }
                else if (
                    string.Equals(
                        payload.Edition,
                        "Enterprise",
                        StringComparison.OrdinalIgnoreCase))
                {
                    licenseType =
                        LicenseType.Enterprise;
                }
                else
                {
                    return null;
                }


                // =================================================
                // PARSE EXPIRY
                // =================================================

                if (!DateTime.TryParseExact(
                        payload.ExpiryDate,
                        "yyyy-MM-dd",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime expiryDate))
                {
                    return null;
                }


                expiryDate =
                    expiryDate.Date;


                LicenseStatus status =
                    DateTime.Today <= expiryDate
                        ? LicenseStatus.Active
                        : LicenseStatus.Expired;


                // =================================================
                // RESULT
                // =================================================

                return
                    new LicenseInfo
                    {
                        Type =
                            licenseType,

                        Status =
                            status,

                        LicensedTo =
                            payload.Customer,

                        ExpiryDate =
                            expiryDate
                    };
            }
            catch
            {
                return null;
            }
        }


        // =========================================================
        // LOAD PUBLIC KEY FROM EMBEDDED RESOURCE
        // =========================================================

        private static string LoadPublicKey()
        {
            Assembly assembly =
                Assembly.GetExecutingAssembly();


            using Stream? stream =
                assembly.GetManifestResourceStream(
                    PublicKeyResourceName);


            if (stream == null)
            {
                throw new InvalidOperationException(
                    "Embedded licensing public key was not found.");
            }


            using StreamReader reader =
                new StreamReader(
                    stream);


            return
                reader.ReadToEnd();
        }


        // =========================================================
        // BASE64 URL DECODE
        // =========================================================

        private static byte[] Base64UrlDecode(
            string value)
        {
            string base64 =
                value
                    .Replace(
                        '-',
                        '+')
                    .Replace(
                        '_',
                        '/');


            switch (base64.Length % 4)
            {
                case 2:

                    base64 +=
                        "==";

                    break;


                case 3:

                    base64 +=
                        "=";

                    break;
            }


            return
                Convert.FromBase64String(
                    base64);
        }


        // =========================================================
        // PAYLOAD
        // =========================================================

        private sealed class LicensePayload
        {
            public string LicenseId { get; set; } =
                "";


            public string Customer { get; set; } =
                "";


            public string Edition { get; set; } =
                "";


            public string IssuedDate { get; set; } =
                "";


            public string ExpiryDate { get; set; } =
                "";
        }
    }
}