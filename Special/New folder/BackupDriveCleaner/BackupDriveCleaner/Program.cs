using System;
using System.Collections.Generic;
using System.Text;

namespace BackupDriveCleaner
{
    class Program
    {

        public const long BackupSizeInBytes = 60000000000;


        public struct BackupDetails
        {
            public string FolderPath;
            public long FolderSize;

        }
        public static System.IO.DriveInfo[] SystemDrives;
        public static List<BackupDetails> BackupLocations = new List<BackupDetails>();
        static void Main(string[] args)
        {

            Console.Clear();
            Console.WriteLine("======================================================");
            Console.WriteLine("=      BackupDriveCleaner Utility   8-28-2014        =");
            Console.WriteLine("=                       version 0.0.1                =");
            Console.WriteLine("======================================================\r\n\r\n");
            
                SystemDrives = System.IO.DriveInfo.GetDrives();
            
            

            bool foundDrive = false;
            Object bd = new Object();
            try
            {
                foreach (System.IO.DriveInfo drive in SystemDrives)
                {
                    if (drive.IsReady == true)
                    {
                        if (drive.VolumeLabel.ToUpper() == "FREEAGENT DRIVE")
                        {
                            //this is the correct drive...
                            Console.WriteLine("Found the backup drive...");
                            foundDrive = true;
                            bd = drive;
                            break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("A drive was found offline.");
                    }
                }
            }
            catch (Exception wtf)
            {
                Console.WriteLine("Exception: " + wtf.Message);
                System.Threading.Thread.Sleep(1000);
                Environment.Exit(0);
            }
           
            if (foundDrive == true)
            {
                Console.WriteLine("Now we would check the drive for available space...");
                System.IO.DriveInfo BackupDrive = (System.IO.DriveInfo)bd;
                Console.WriteLine("Backup Drive is " + ((decimal)((decimal)(BackupDrive.TotalSize - BackupDrive.TotalFreeSpace) / BackupDrive.TotalSize) * 100).ToString("0.##") + "% full.");
                Console.WriteLine("Free Space on drive (" + BackupDrive.RootDirectory.ToString() + "): " + BackupDrive.TotalFreeSpace.ToString() + " bytes");
                Console.WriteLine("Backup Space takes about:   " + BackupSizeInBytes.ToString() + " bytes");
                if (BackupDrive.TotalFreeSpace < BackupSizeInBytes)
                {
                    Console.WriteLine("Going to delete the oldest backups to make enough free space on drive");

                    long SpaceToShave = BackupSizeInBytes - BackupDrive.TotalFreeSpace;
                    Console.WriteLine("Space to shave: " + SpaceToShave.ToString() + " bytes");
                    string[] Folders = System.IO.Directory.GetDirectories(BackupDrive.RootDirectory.ToString());
                    bool hasBEFolder = false;
                    foreach (string fld in Folders)
                    {
                        //Console.WriteLine("Checking folder " + fld);
                        if (fld.ToUpper().Split('\\')[1] == "BE")
                        {
                            hasBEFolder = true;
                            break;
                        }
                    }

                    if (hasBEFolder == true)
                    {
                        Console.WriteLine("Found the 'BE' folder. Checking inside...");
                        string BEDirectory = BackupDrive.RootDirectory.ToString() + "BE\\";
                        Console.WriteLine("Checking directory " + BEDirectory);
                        string[] BEFolders = System.IO.Directory.GetDirectories(BEDirectory);
                        
                        
                        List<string> BackupFolders = new List<string>();

                        foreach (string fld in BEFolders)
                        {                            
                            if (fld.ToUpper().Split('\\')[2].Substring(0, 3) == "IMG")
                            {
                                BackupFolders.Add(fld);
                            }
                        }

                        
                        Console.WriteLine("Found Backups-\r\n================================================");
                        foreach (string bf in BackupFolders)
                        {
                            BackupDetails bde = new BackupDetails();
                            bde.FolderPath = bf;
                            bde.FolderSize = GetFolderSize(bf);
                            BackupLocations.Add(bde);

                            System.IO.DirectoryInfo DI = new System.IO.DirectoryInfo(bf);
                            Console.WriteLine(bf + "  -  " + bde.FolderSize.ToString() + " bytes");
                        }
                        Console.WriteLine("================================================\r\n");

                        Console.WriteLine("Removal List:\r\n==============================");

                        int bckCount = 0;
                        long spaceSaved = 0;
                        while (BackupDrive.TotalFreeSpace + spaceSaved < BackupSizeInBytes)
                        {
                            try
                            {
                                System.Threading.Thread.Sleep(200);
                                Console.Write("Removing Backup " + BackupLocations[bckCount].FolderPath + "...");
                                //REMOVE THE BACKUP...//////////////////////////////
                                System.IO.Directory.Delete(BackupLocations[bckCount].FolderPath, true);
                                ////////////////////////////////////////////////////
                                spaceSaved += BackupLocations[bckCount].FolderSize;
                                System.Threading.Thread.Sleep(500);
                                Console.Write(" OK.\r\n");                                
                                
                                bckCount++;
                            }
                            catch
                            {
                                Console.WriteLine("There is not enough free space to make room for a backup.\r\n");
                                break;
                            }
                        }

                        Console.WriteLine("==============================\r\n");

                        Console.WriteLine("The BackupDriveCleaner has finished processing.");
                        Console.WriteLine("You removed " + (spaceSaved / 1000000000) + " GB of older data.");
                        Console.WriteLine("Estimated free space after next backup is " + ((BackupDrive.TotalFreeSpace / 1000000000) + ((spaceSaved / 1000000000) - (BackupSizeInBytes / 1000000000))).ToString() + " GB.");
                        Console.WriteLine("See you later!");
                        System.Threading.Thread.Sleep(10000);
                        
                        Environment.Exit(0);

                    }
                    else
                    {
                        Console.WriteLine("Missing 'BE' folder. Exiting.");
                        System.Threading.Thread.Sleep(10000);
                        Environment.Exit(0);
                    }



                }
                else
                {
                    Console.WriteLine("This drive has enough free space for a backup run.");
                    System.Threading.Thread.Sleep(10000);
                    Environment.Exit(0);
                }
                
                Console.Clear();
            }
            else
            {
                Console.WriteLine("Sorry, the drive was not found...");
                System.Threading.Thread.Sleep(10000);
                Environment.Exit(0);
            }




        }


        private static long GetFolderSize(string folderPath)
        {

            string[] dirs = System.IO.Directory.GetDirectories(folderPath);
            long ByteCount = 0;
            foreach (string directory in dirs)
            {
                foreach (string file in System.IO.Directory.GetFiles(directory))
                {
                    ByteCount += new System.IO.FileInfo(file).Length;
                }
            }
            foreach (string file in System.IO.Directory.GetFiles(folderPath))
            {
                ByteCount += new System.IO.FileInfo(file).Length;
            }

            return ByteCount;
        }
    }
}
