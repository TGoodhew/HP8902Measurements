using Ivi.Visa;
using NationalInstruments.Visa;
using Spectre.Console;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using System.IO;
using System.Net.NetworkInformation;

namespace HP8902Measurements
{
    public class CalibrationFactor
    {
        public decimal Frequency { get; set; }
        public decimal CalFactor { get; set; }

        public CalibrationFactor() { }

        public CalibrationFactor(double frequency, double calFactor)
        {
            Frequency = (decimal)frequency;
            CalFactor = (decimal)calFactor;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            int gpibAddress8902A = 14; // This is the default address for an 8902A
            int gpibAddress8673B = 19; // This is the default address for the 8673B
            SemaphoreSlim srqWait = new SemaphoreSlim(0, 1);
            NationalInstruments.Visa.ResourceManager resManager = null;
            GpibSession gpibSession8902A = null;
            GpibSession gpibSession8673B = null;

            DisplayeTitle(gpibAddress8902A, gpibAddress8673B);

            // Ask for test choice
            var TestChoice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Select the test to run?")
                    .PageSize(10)
                    .AddChoices(new[] { "Set GPIB Addresses", "Connect to instruments", "Calibrate Sensor","Set expected frequency", "Exit" })
                    );

            while (TestChoice != "Exit")
            {
                switch (TestChoice)
                {
                    case "Set GPIB Addresses":
                        SetGPIBAddresses(ref gpibAddress8902A, ref gpibAddress8673B);
                        break;
                    case "Connect to instruments":
                        {
                            // Connect to 8902A
                            if (!ConnectToDevice(gpibAddress8902A, ref srqWait, ref resManager, ref gpibSession8902A))
                            {
                                break;
                            }
                            // Connect to 8673B
                            if (!ConnectToDevice(gpibAddress8673B, ref srqWait, ref resManager, ref gpibSession8673B))
                            {
                                break;
                            }
                            break;
                        }

                    case "Calibrate Sensor":
                        {
                            if (gpibSession8902A == null || gpibSession8673B == null)
                            {
                                AnsiConsole.MarkupLine("[red]Error: Both instruments must be connected before calibration can proceed.[/]");
                                Thread.Sleep(1000); // Pause for a moment to let the user see the message
                                break;
                            }
                            // Start calibration process
                            AnsiConsole.MarkupLine("[green]Starting calibration process...[/]");

                            // Load calibration factors
                            var calibrationFactors = LoadCalibrationTableFile(gpibSession8902A);

                            var confirmation = AnsiConsole.Prompt(new ConfirmationPrompt("Sensor must be connected to the RF Power Output to continue. Is it connected?"));

                            // If not connected abort calibration
                            // TODO: This could be improved with a check of the sensor connection status
                            if (!confirmation)
                            {
                                AnsiConsole.MarkupLine("[red]Calibration aborted by user.[/]");
                                Thread.Sleep(1000); // Pause for a moment to let the user see the message
                                break;
                            }

                            // Zero the sensor
                            AnsiConsole.MarkupLine("[green]Zeroing the sensor...[/]");

                            // Place the unit into RF Power mode and with the trigger off (free run)
                            SendCommand("M4T0", gpibSession8902A);

                            // Zero the unit
                            SendCommand("ZR", gpibSession8902A);

                            // When a Zero command is sent we should wait for a data item to be ready to confirm that the Zero has been completed (SP22.3)
                            SendCommand("22.3SP", gpibSession8902A);

                            // Wait for the data to be available
                            srqWait.Wait();

                            // Clear the SRQ Mask
                            SendCommand("22.0SP", gpibSession8902A);

                            // Calibrate the sesnor
                            AnsiConsole.MarkupLine("[green]Calibrating the sensor...[/]");

                            // Place the unit into RF Power mode and with the trigger off (free run)
                            SendCommand("M4T0", gpibSession8902A);

                            // Turn the calibration source on
                            SendCommand("C1", gpibSession8902A);

                            // Wait for calibration to complete
                            SendCommand("22.3SP", gpibSession8902A);

                            // Wait for the data to be available
                            srqWait.Wait();

                            // Clear the SRQ Mask
                            SendCommand("22.0SP", gpibSession8902A);

                            // Save the calibration
                            SendCommand("SC", gpibSession8902A);

                            // Turn the calibration source off
                            SendCommand("C0", gpibSession8902A);

                            AnsiConsole.MarkupLine("[green]Calibration process completed.[/]");
                            Thread.Sleep(1000); // Pause for a moment to let the user see the message
                            break;
                        }

                    case "Set expected frequency":
                        {
                            if (gpibSession8902A == null || gpibSession8673B == null)
                            {
                                AnsiConsole.MarkupLine("[red]Error: Both instruments must be connected before setting frequency can proceed.[/]");
                                Thread.Sleep(1000); // Pause for a moment to let the user see the message
                                break;
                            }
                            // Ask user for frequency
                            double frequencyGHz = AnsiConsole.Prompt(
                                new TextPrompt<double>("Enter the expected frequency in GHz (0.00015 to 18.0 GHz)?")
                                .Validate(n => n >= 0.00015 && n <= 18.0 ? ValidationResult.Success() : ValidationResult.Error("Frequency must be between 0.00015 and 18.0 GHz"))
                                );

                            // If the frequency is 1300MHz or less, then we can turn off the LO setting for this reading.
                            if (frequencyGHz <= 1.3)
                            {
                                SendCommand("27.3SP0MZ", gpibSession8902A); // Disable LO

                                // Set the 8673B output power to a nominal level
                                SendCommand("FR3GZLE-70DM", gpibSession8673B);
                            }
                            else
                            {
                                double[] increments = { 0.12053, 0.24053, 0.48053, 0.60053, 0.68053 };

                                double chosenIncrement = 0;
                                foreach (double inc in increments)
                                {
                                    if (frequencyGHz + inc > 2.0)
                                    {
                                        chosenIncrement = inc;
                                        break; // stop at the first increment that works
                                    }
                                }

                                // Enable the LO on the 8902A with the chosen frequency
                                SendCommand("27.3SP"+((frequencyGHz+ chosenIncrement)*1000).ToString()+"MZ", gpibSession8902A); // Enable LO

                                // Set the 8673B to the correct LO frequency
                                SendCommand(string.Format("FR{0:G}GZ", frequencyGHz + chosenIncrement), gpibSession8673B);

                                // Set the 8673B output power to a nominal level
                                SendCommand("LE8DM", gpibSession8673B);
                            }

                            AnsiConsole.MarkupLine("[green]Expected frequency set on both instruments.[/]");
                            Thread.Sleep(1000); // Pause for a moment to let the user see the message
                            break;
                        }
                }

                // Clear the screen & Display title
                DisplayeTitle(gpibAddress8902A, gpibAddress8673B);

                // Ask for test choice
                TestChoice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Select the test to run?")
                    .PageSize(10)
                    .AddChoices(new[] { "Set GPIB Addresses", "Connect to instruments", "Calibrate Sensor", "Set expected frequency", "Exit" })
                    );
            }

        }

        private static void SetGPIBAddresses(ref int gpibAddress8902A, ref int gpibAddress8673B)
        {
            int localGpibAddress8673B = gpibAddress8673B; // Capture value for lambda

            gpibAddress8902A = AnsiConsole.Prompt(
                new TextPrompt<int>("Enter HP 8902A GPIB address (Default is 14)?")
                .DefaultValue(14)
                .Validate(n => n >= 1 && n <= 30 ? ValidationResult.Success() : ValidationResult.Error("Address must be between 1 and 30"))
                .Validate(n => n != localGpibAddress8673B ? ValidationResult.Success() : ValidationResult.Error("8902A and 8673B addresses must be different"))
                );

            int localGpibAddress8902A = gpibAddress8902A; // Capture value for lambda
            gpibAddress8673B = AnsiConsole.Prompt(
                new TextPrompt<int>("Enter HP 8673B GPIB address (Default is 19)?")
                .DefaultValue(19)
                .Validate(n => n >= 1 && n <= 30 ? ValidationResult.Success() : ValidationResult.Error("Address must be between 1 and 30"))
                .Validate(n => n != localGpibAddress8902A ? ValidationResult.Success() : ValidationResult.Error("8902A and 8673B addresses must be different"))
                );

            AnsiConsole.MarkupLine("[green]GPIB Addresses updated.[/]");
            Thread.Sleep(1000); // Pause for a moment to let the user see the message

            return;
        }

        private static bool ConnectToDevice(int gpibIntAddress, ref SemaphoreSlim srqWait, ref ResourceManager resManager, ref GpibSession gpibSession)
        {
            if (resManager == null)
            {
                resManager = new NationalInstruments.Visa.ResourceManager();
            }

            if (gpibSession != null)
            {
                AnsiConsole.MarkupLine("[yellow]Warning: Already connected to a device. Disconnecting and reconnecting.[/]");
                gpibSession?.Dispose();
                resManager?.Dispose();
                resManager = null;
                gpibSession = null;
                Thread.Sleep(1000); // Pause for a moment to let the user see the message
            }

            try
            {
                // Setup the GPIB connection via the ResourceManager
                resManager = new NationalInstruments.Visa.ResourceManager();

                // Create a GPIB session for the specified address
                gpibSession = (GpibSession)resManager.Open(string.Format("GPIB0::{0}::INSTR", gpibIntAddress));
                gpibSession.TimeoutMilliseconds = 2000; // Set the timeout to be 2s
                gpibSession.TerminationCharacterEnabled = true;
                gpibSession.Clear(); // Clear the session

                var sessionCopy = gpibSession;
                var srqWaitCopy = srqWait;
                gpibSession.ServiceRequest += (sender, e) => SRQHandler(sender, e, sessionCopy, srqWaitCopy);

                var commandSent = SendCommand("IP", gpibSession);

                if (!commandSent)
                {
                    AnsiConsole.MarkupLine("[Red]Device failed to connect. Check GPIB address and device state.[/]");
                    Thread.Sleep(1000); // Pause for a moment to let the user see the message
                    resManager = null;
                    gpibSession = null;
                    return false;
                }
                else
                {
                    // Successfully connected
                    AnsiConsole.MarkupLine("[green]Device connected: [/]" + gpibIntAddress);
                    Thread.Sleep(1000); // Pause for a moment to let the user see the message
                    return true;
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[Red]Device failed to connect. GPIB Error: {ex.Message}[/]");
                Thread.Sleep(1000); // Pause for a moment to let the user see the message
                resManager = null;
                gpibSession = null;
                return false;
            }
        }

        private static List<CalibrationFactor> LoadCalibrationTableFile(GpibSession gpibSession8902A)
        {
            string fileName = "CalFactors92A.json";
            if (!File.Exists(fileName))
            {
                CreateCalibrationTableFile();
                AnsiConsole.MarkupLine("[yellow]Warning: Loading default calibration factor file.[/]");
            }
            string jsonString = File.ReadAllText(fileName);
            List<CalibrationFactor> calibrationFactors = JsonSerializer.Deserialize<List<CalibrationFactor>>(jsonString);

            // Set frequency offset table mode as we're using the down converter
            SendCommand("27.1SP", gpibSession8902A);

            // Clear the table
            SendCommand("37.9SP", gpibSession8902A);

            // Store the calibration factors
            foreach (CalibrationFactor factor in calibrationFactors)
            {
                SendCommand("37.3SP" + String.Format("{0:F2}MZ{1:F2}CF", factor.Frequency * 1000, factor.CalFactor), gpibSession8902A);
            }

            // Successfully connected
            AnsiConsole.MarkupLine("[green]Calibration factors loaded.[/]");
            Thread.Sleep(1000); // Pause for a moment to let the user see the message

            return calibrationFactors;
        }

        private static void CreateCalibrationTableFile()
        {
            // HP 11792A Table
            List<CalibrationFactor> calibrationFactors92A = new List<CalibrationFactor>();

            calibrationFactors92A.Add(new CalibrationFactor(0.05, 100.0));
            calibrationFactors92A.Add(new CalibrationFactor(2.0, 96.3));
            calibrationFactors92A.Add(new CalibrationFactor(3.0, 94.8));
            calibrationFactors92A.Add(new CalibrationFactor(4.0, 93.9));
            calibrationFactors92A.Add(new CalibrationFactor(5.0, 92.9));
            calibrationFactors92A.Add(new CalibrationFactor(6.0, 91.9));
            calibrationFactors92A.Add(new CalibrationFactor(7.0, 91.1));
            calibrationFactors92A.Add(new CalibrationFactor(8.0, 90.3));
            calibrationFactors92A.Add(new CalibrationFactor(9.0, 89.3));
            calibrationFactors92A.Add(new CalibrationFactor(10.0, 88.5));
            calibrationFactors92A.Add(new CalibrationFactor(11.0, 87.5));
            calibrationFactors92A.Add(new CalibrationFactor(12.4, 87.0));
            calibrationFactors92A.Add(new CalibrationFactor(13.0, 86.1));
            calibrationFactors92A.Add(new CalibrationFactor(14.0, 85.6));
            calibrationFactors92A.Add(new CalibrationFactor(15.0, 85.4));
            calibrationFactors92A.Add(new CalibrationFactor(16.0, 84.9));
            calibrationFactors92A.Add(new CalibrationFactor(17.0, 84.6));
            calibrationFactors92A.Add(new CalibrationFactor(18.0, 84.1));

            AnsiConsole.MarkupLine("[yellow]Warning: Creating default calibration factor file.[/]");

            string fileName = "CalFactors92A.json";
            string jsonString = JsonSerializer.Serialize(calibrationFactors92A);
            File.WriteAllText(fileName, jsonString);
        }

        private static void DisplayeTitle(int gpibAAddress8902A, int gpibAAddress8673B)
        {
            // Clear screen and display header
            AnsiConsole.Clear();
            AnsiConsole.Write(
                new FigletText("HP8902A Mesurements")
                    .LeftJustified()
                    .Color(Color.Green));
            AnsiConsole.WriteLine("--------------------------------------------------");
            AnsiConsole.WriteLine("");
            AnsiConsole.WriteLine("HP8902A Mesurements - Simple UI for conducting measurement up to 18GHz");
            AnsiConsole.WriteLine("");
            AnsiConsole.WriteLine("HP 8902A GPIB Address: " + gpibAAddress8902A);
            AnsiConsole.WriteLine("HP 8673B GPIB Address: " + gpibAAddress8673B);
            AnsiConsole.WriteLine("");
        }

        static private bool SendCommand(string command, GpibSession gpibSession)
        {
            try
            {
                gpibSession.FormattedIO.WriteLine(command);
                return true;
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]GPIB Send Error: {ex.Message}[/]");
                Debug.WriteLine($"GPIB Send Error: {ex}");
                return false;
            }
        }

        static private string ReadResponse(GpibSession gpibSession)
        {
            try
            {
                return gpibSession.FormattedIO.ReadLine();
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]GPIB Read Error: {ex.Message}[/]");
                Debug.WriteLine($"GPIB Read Error: {ex}");
                return string.Empty;
            }
        }
        static private string QueryString(string command, GpibSession gpibSession)
        {
            SendCommand(command, gpibSession);
            var response = ReadResponse(gpibSession);
            if (string.IsNullOrWhiteSpace(response))
            {
                AnsiConsole.MarkupLine("[yellow]Warning: No response from instrument.[/]");
            }
            return response;
        }

        public static void SRQHandler(object sender, Ivi.Visa.VisaEventArgs e, GpibSession gpibSession, SemaphoreSlim srqWait)
        {
            try
            {
                var gbs = (GpibSession)sender;
                StatusByteFlags sb = gbs.ReadStatusByte();

                Debug.WriteLine($"SRQHandler - Status Byte: {sb}");

                gpibSession.DiscardEvents(EventType.ServiceRequest);

                SendCommand("*CLS", gpibSession);

                srqWait.Release();
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]SRQ Handler Error: {ex.Message}[/]");
                Debug.WriteLine($"SRQ Handler Error: {ex}");
            }
        }

    }
}
