using ICM300_1182PNC;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ICRWActiveX
{
    [ProgId("ICRWActiveX.CardMSData")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("5BF81A8B-5600-4FF5-8C50-6FD45981905F")]
    [ComVisible(true)]

    public class CardMSData
    {
        IDCDeviceCtrl deviceCtrl = new IDCDeviceCtrl();
        IDCCtrl.EnumCmdResult result;
        short trackDataPresence;
        short cardDetected;


        string Initialize()
        // Executes Initialize command
        {
            Byte[] command = new Byte[] { 0x43, 0x30, 0x30, };
            Int32 Timeout = 20000;             // milliseconds
            Byte[] response;

            // Executes command, and then receives a reply for the command
            result = deviceCtrl.ExecuteCommand(
                                    command,
                                    Timeout,
                                    out response
                                    );

            if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
            {
                //Console.WriteLine(
                //                "ExecuteCommand failed returning the error '{0}'.\n",
                //                result.ToString()
                //                );

                // Closes communications
                deviceCtrl.DisconnectDevice();

                return "Init Failed - Initialize Command failed";
            }
            else
            {
                if (response[0] == 0x50)
                {
                    // Received positive response
                    //Console.WriteLine("Initialize command succeeded.\n");
                    return "Initialize command successful.";
                }
                else if (response[0] == 0x4E)
                {
                    //Console.WriteLine(
                    //                "Initialize command failed returning the error code '{0}{1}'.\n",
                    //                (char)response[3],
                    //                (char)response[4]
                    //                );

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return "Initialize command failed returning the error code";
                }
                else
                {
                    // Received unexpected response
                    //Console.WriteLine("Unexpected error occurred when executing Initialize command.\n");

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return "Unexpected error occurred when executing Initialize command.";
                }
            }
        }

        string ReadTrack1()
        // Read Track1 Data
        {
            Byte[] command = new Byte[] { 0x43, 0x32, 0x31, };
            Int32 Timeout = 20000;             // milliseconds
            Byte[] response;

            // Executes command, and then receives a reply for the command
            result = deviceCtrl.ExecuteCommand(
                                    command,
                                    Timeout,
                                    out response
                                    );

            if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
            {
                //Console.WriteLine(
                //                "ExecuteCommand failed returning the error '{0}'.\n",
                //                result.ToString()
                //                );

                // Closes communications
                deviceCtrl.DisconnectDevice();

                return null;
            }
            else
            {
                if (response[0] == 0x50)
                {
                    // Received positive response
                    //Console.WriteLine("ReadTrack1 command succeeded.\n");
                    string trackData = System.Text.Encoding.UTF8.GetString(response, (int)5, response.Length - (int)5);
                    //Console.WriteLine("Track 1 Data - {0} \n", trackData);
                    string[] trackParts = trackData.Split('^');
                    //Console.WriteLine("Card Number - {0} \n", trackParts[0].Substring(1));
                    //Console.WriteLine("Customer Name - {0} \n", trackParts[1]);
                    //return trackParts[0].Substring(1);
                    return trackData;
                }
                else if (response[0] == 0x4E)
                {
                    //Console.WriteLine(
                    //                "ReadTrack1 command failed returning the error code '{0}{1}'.\n",
                    //                (char)response[3],
                    //                (char)response[4]
                    //                );

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return null;
                }
                else
                {
                    // Received unexpected response
                    //Console.WriteLine("Unexpected error occurred when executing Read Track1 command.\n");

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return null;
                }
            }
        }

        string ReadTrack2()
        // Executes ReadTrack2 command
        {
            Byte[] command = new Byte[] { 0x43, 0x32, 0x32, };
            Int32 Timeout = 20000;             // milliseconds
            Byte[] response;

            // Executes command, and then receives a reply for the command
            result = deviceCtrl.ExecuteCommand(
                                    command,
                                    Timeout,
                                    out response
                                    );

            if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
            {
                //Console.WriteLine(
                //                "ExecuteCommand failed returning the error '{0}'.\n",
                //                result.ToString()
                //                );

                // Closes communications
                deviceCtrl.DisconnectDevice();

                return null;
            }
            else
            {
                if (response[0] == 0x50)
                {
                    // Received positive response
                    //Console.WriteLine("ReadTrack2 command succeeded.\n");
                    string trackData = System.Text.Encoding.UTF8.GetString(response, (int)5, response.Length - (int)5);
                    //Console.WriteLine("Track 2 Data - {0} \n", trackData);
                    return trackData;
                }
                else if (response[0] == 0x4E)
                {
                    //Console.WriteLine(
                    //                "ReadTrack2 command failed returning the error code '{0}{1}'.\n",
                    //                (char)response[3],
                    //                (char)response[4]
                    //                );

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return null;
                }
                else
                {
                    // Received unexpected response
                    //Console.WriteLine("Unexpected error occurred when executing ReadTrack2 command.\n");

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return null;
                }
            }
        }

        string ReadTrack3()
        // Executes ReadTrack3 command
        {
            Byte[] command = new Byte[] { 0x43, 0x32, 0x33, };
            Int32 Timeout = 20000;             // milliseconds
            Byte[] response;

            // Executes command, and then receives a reply for the command
            result = deviceCtrl.ExecuteCommand(
                                    command,
                                    Timeout,
                                    out response
                                    );

            if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
            {
                //Console.WriteLine(
                //                "ExecuteCommand failed returning the error '{0}'.\n",
                //                result.ToString()
                //                );

                // Closes communications
                //deviceCtrl.DisconnectDevice();

                return null;
            }
            else
            {
                if (response[0] == 0x50)
                {
                    // Received positive response
                    //Console.WriteLine("ReadTrack3 command succeeded.\n");
                    string trackData = System.Text.Encoding.UTF8.GetString(response, (int)5, response.Length - (int)5);
                    //Console.WriteLine("Track 3 Data - {0} \n", trackData);
                    return trackData;
                }
                else if (response[0] == 0x4E)
                {
                    //Console.WriteLine(
                    //                "ReadTrack3 command failed returning the error code '{0}{1}'.\n",
                    //                (char)response[3],
                    //                (char)response[4]
                    //                );

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return null;
                }
                else
                {
                    // Received unexpected response
                    //Console.WriteLine("Unexpected error occurred when executing ReadTrack3 command.\n");

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return null;
                }
            }
        }

        void Status()
        // Executes Status command
        {
            Byte[] command = new Byte[] { 0x43, 0x31, 0x30, };
            Int32 Timeout = 20000;             // milliseconds
            Byte[] response;

            // Executes command, and then receives a reply for the command
            result = deviceCtrl.ExecuteCommand(
                                    command,
                                    Timeout,
                                    out response
                                    );

            if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
            {
                Console.WriteLine(
                                "ExecuteCommand failed returning the error '{0}'.\n",
                                result.ToString()
                                );

                // Closes communications
                deviceCtrl.DisconnectDevice();

                return;
            }
            else
            {
                if (response[0] == 0x50)
                {
                    // Received positive response
                    Console.WriteLine("Status command succeeded.\n");
                    Func(response);
                }
                else if (response[0] == 0x4E)
                {
                    Console.WriteLine(
                                    "Status command failed returning the error code '{0}{1}'.\n",
                                    (char)response[3],
                                    (char)response[4]
                                    );

                    // Closes communications
                    //deviceCtrl.DisconnectDevice();

                    return;
                }
                else
                {
                    // Received unexpected response
                    Console.WriteLine("Unexpected error occurred when executing Status command.\n");

                    // Closes communications
                    deviceCtrl.DisconnectDevice();

                    return;
                }
            }
        }

        void Func(Byte[] bufferStatusReport)
        {
            trackDataPresence = 0;
            cardDetected = 0;
            if (bufferStatusReport[0] == 0x50 || bufferStatusReport[0] == 0x52)
            {
                // Received positive response
                Console.WriteLine("Positive Status Received.\n");
            }
            else if (bufferStatusReport[0] == 0x4E)
            {
                Console.WriteLine(
                                "Initialize command failed returning the error code '{0}{1}'.\n",
                                (char)bufferStatusReport[3],
                                (char)bufferStatusReport[4]
                                );
            }
            if (bufferStatusReport[3] == 0x30 && bufferStatusReport[4] == 0x30)
            {
                Console.WriteLine(
                                "No Card Inside. Ready to accept card .\n");
            }
            else
            {
                if (bufferStatusReport[3] == 0x31 || bufferStatusReport[3] == 0x32 || bufferStatusReport[3] == 0x33)
                {
                    //Console.WriteLine("Card Detected Inside.\n");
                    cardDetected = 1;
                }
                switch (bufferStatusReport[4])
                {
                    case 0x30:
                        //Console.WriteLine(
                        //           "Please try once again .\n");
                        //Initialize();
                        trackDataPresence = 0;
                        break;
                    case 0x31:
                        //Console.WriteLine(
                        //           "Card has Track 1 Data - .\n");
                        trackDataPresence = 1;
                        break;
                    case 0x32:
                        //Console.WriteLine(
                        //           "Card has Track 2 Data - .\n");
                        trackDataPresence = 2;
                        break;
                    case 0x33:
                        //Console.WriteLine(
                        //           "Card has both Track 1 &  Track 2 Data - .\n");
                        trackDataPresence = 3;
                        break;
                    case 0x34:
                        //Console.WriteLine(
                        //           "Card has Track 3 Data - .\n");
                        trackDataPresence = 4;
                        break;
                    case 0x35:
                        //Console.WriteLine(
                        //           "Card has both Track 1 &  Track 3 Data - .\n");
                        trackDataPresence = 5;
                        break;
                    case 0x36:
                        //Console.WriteLine(
                        //           "Card has both Track 2 &  Track 3 Data - .\n");
                        trackDataPresence = 6;
                        break;
                    case 0x37:
                        //Console.WriteLine(
                        //           "Card has all Track 1, Track 2 & Track 3 Data - .\n");
                        trackDataPresence = 7;
                        break;
                }
            }
        }

        [ComVisible(true)]
        public string Init(string com, string baud, bool mock)
        {
            if (mock)
            {
                return "Init Successful.";
            }
            String ComPortNumber = com;
            Int32 Baudrate = Convert.ToInt32(baud);

            // Establishes communications between the Host Computer and the Card Reader/Writer
            {
                result = deviceCtrl.ConnectDevice(ComPortNumber, Baudrate);
                if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
                {
                    return "Init Failed - Connection couldn't be established.";
                }

                result = deviceCtrl.SetStatusReportFunction(Func);
                if (result != IDCCtrl.EnumCmdResult._NO_ERROR)
                {
                    return "Init Failed - SetStatusReport Failed.";
                }
            }
            string temp = Initialize();
            if (temp.Equals("Initialize command successful."))
                return "Init Successful.";
            else
                return temp;

        }


        [ComVisible(true)]
        public void Disconnect()
        {
            // Closes communications
            deviceCtrl.DisconnectDevice();
        }

        [ComVisible(true)]
        public GetDataEntities GetTrackData(bool mock)
        {
            GetDataEntities getDataEntities = new GetDataEntities();
            string output;
            if (mock)
            {
                getDataEntities.Track1 = "Your card is hacked";
                getDataEntities.Track2 = "Now I know your secret key too";
                getDataEntities.Output = "Card Detected";
                return getDataEntities;
            }
            
            //Console.WriteLine("Please Insert Card \n");
            if (cardDetected == 1)
                getDataEntities.Output = "Card Detected";
            else
                output = "No Card Detected";
            if (trackDataPresence == 0)
                getDataEntities.Output += "No Track Data Found";
            getDataEntities.Track1 = getDataEntities.Track1 = "";

            if ((trackDataPresence == 1) || (trackDataPresence == 3) || (trackDataPresence == 5))
            {
                getDataEntities.Track1 = ReadTrack1();
            }
            if ((trackDataPresence == 2) || (trackDataPresence == 3) || (trackDataPresence == 6))
            {
                getDataEntities.Track2 = ReadTrack2();
            }

            return getDataEntities;
        }
    }
}
