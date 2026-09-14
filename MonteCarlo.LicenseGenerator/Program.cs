using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MonteCarlo.LicenseGenerator
{
    internal class Program
    {
        private const string LicensePrefix =
            "MC1";


        static void Main(
            string[] args)
        {
            Console.WriteLine(
                "Monte Carlo Customer License Generator");

            Console.WriteLine(
                "=====================================");

            Console.WriteLine();


            // =====================================================
            // LOCATE PRIVATE KEY
            // =====================================================

            string keyFolder =
                Path.Combine(
                    AppContext.BaseDirectory,
                    "Keys");


            string privateKeyPath =
                Path.Combine(
                    keyFolder,
                    "private-key.pem");


            if (!File.Exists(
                    privateKeyPath))
            {
                Console.WriteLine(
                    "ERROR: private-key.pem was not found.");

                Console.WriteLine();

                Console.WriteLine(
                    $"Expected location:\n{privateKeyPath}");

                Console.WriteLine();

                Console.WriteLine(
                    "Copy your private key into the Keys folder " +
                    "for this generator only.");

                return;
            }


            // =====================================================
            // CUSTOMER NAME
            // =====================================================

            Console.Write(
                "Customer name: ");


            string customerName =
                Console.ReadLine()?
                    .Trim()
                ?? "";


            if (string.IsNullOrWhiteSpace(
                    customerName))
            {
                Console.WriteLine(
                    "Customer name is required.");

                return;
            }


            // =====================================================
            // EDITION
            // =====================================================

            Console.WriteLine();

            Console.WriteLine(
                "Edition");

            Console.WriteLine(
                "1 = Professional");

            Console.WriteLine(
                "2 = Enterprise");

            Console.Write(
                "Select edition: ");


            string editionInput =
                Console.ReadLine()?
                    .Trim()
                ?? "";


            string edition;


            switch (editionInput)
            {
                case "1":

                    edition =
                        "Professional";

                    break;


                case "2":

                    edition =
                        "Enterprise";

                    break;


                default:

                    Console.WriteLine(
                        "Invalid edition.");

                    return;
            }


            // =====================================================
            // EXPIRY DATE
            // =====================================================

            Console.WriteLine();

            Console.Write(
                "Expiry date (dd-MM-yyyy): ");


            string expiryInput =
                Console.ReadLine()?
                    .Trim()
                ?? "";


            if (!DateTime.TryParseExact(
                    expiryInput,
                    "dd-MM-yyyy",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out DateTime expiryDate))
            {
                Console.WriteLine(
                    "Invalid expiry date.");

                return;
            }


            expiryDate =
                expiryDate.Date;


            if (expiryDate < DateTime.Today)
            {
                Console.WriteLine(
                    "Expiry date cannot be in the past.");

                return;
            }


            // =====================================================
            // CREATE LICENSE PAYLOAD
            // =====================================================

            string licenseId =
                "MC-" +
                Guid.NewGuid()
                    .ToString("N")
                    .Substring(
                        0,
                        12)
                    .ToUpperInvariant();


            DateTime issuedDate =
                DateTime.Today;


            LicensePayload payload =
                new LicensePayload
                {
                    LicenseId =
                        licenseId,

                    Customer =
                        customerName,

                    Edition =
                        edition,

                    IssuedDate =
                        issuedDate.ToString(
                            "yyyy-MM-dd"),

                    ExpiryDate =
                        expiryDate.ToString(
                            "yyyy-MM-dd")
                };


            string payloadJson =
                JsonSerializer.Serialize(
                    payload);


            byte[] payloadBytes =
                Encoding.UTF8.GetBytes(
                    payloadJson);


            // =====================================================
            // SIGN PAYLOAD
            // =====================================================

            string privateKeyPem =
                File.ReadAllText(
                    privateKeyPath);


            byte[] signatureBytes;


            using (
                RSA rsa =
                    RSA.Create())
            {
                rsa.ImportFromPem(
                    privateKeyPem);


                signatureBytes =
                    rsa.SignData(
                        payloadBytes,
                        HashAlgorithmName.SHA256,
                        RSASignaturePadding.Pkcs1);
            }


            // =====================================================
            // CREATE LICENSE CODE
            // =====================================================

            string payloadPart =
                Base64UrlEncode(
                    payloadBytes);


            string signaturePart =
                Base64UrlEncode(
                    signatureBytes);


            string licenseCode =
                LicensePrefix +
                "." +
                payloadPart +
                "." +
                signaturePart;


            // =====================================================
            // SAVE OUTPUT
            // =====================================================

            string outputFolder =
                Path.Combine(
                    AppContext.BaseDirectory,
                    "Licenses");


            Directory.CreateDirectory(
                outputFolder);


            string safeCustomerName =
                MakeSafeFileName(
                    customerName);


            string outputFile =
                Path.Combine(
                    outputFolder,
                    $"{safeCustomerName}_" +
                    $"{licenseId}.txt");


            File.WriteAllText(
                outputFile,
                licenseCode);


            // =====================================================
            // DISPLAY RESULT
            // =====================================================

            Console.WriteLine();
            Console.WriteLine(
                "License generated successfully.");

            Console.WriteLine();

            Console.WriteLine(
                $"License ID : {licenseId}");

            Console.WriteLine(
                $"Customer   : {customerName}");

            Console.WriteLine(
                $"Edition    : {edition}");

            Console.WriteLine(
                $"Issued     : {issuedDate:dd-MMM-yyyy}");

            Console.WriteLine(
                $"Expires    : {expiryDate:dd-MMM-yyyy}");

            Console.WriteLine();

            Console.WriteLine(
                $"Saved to:\n{outputFile}");

            Console.WriteLine();

            Console.WriteLine(
                "License code:");

            Console.WriteLine();

            Console.WriteLine(
                licenseCode);
        }


        // =========================================================
        // BASE64 URL ENCODING
        // =========================================================

        private static string Base64UrlEncode(
            byte[] data)
        {
            return
                Convert.ToBase64String(
                    data)
                .TrimEnd('=')
                .Replace(
                    '+',
                    '-')
                .Replace(
                    '/',
                    '_');
        }


        // =========================================================
        // SAFE FILE NAME
        // =========================================================

        private static string MakeSafeFileName(
            string value)
        {
            foreach (
                char invalidChar
                in Path.GetInvalidFileNameChars())
            {
                value =
                    value.Replace(
                        invalidChar,
                        '_');
            }


            return
                value.Replace(
                    ' ',
                    '_');
        }


        // =========================================================
        // LICENSE PAYLOAD
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