using System;
using System.IO;
using System.Management.Automation;
using System.Runtime.InteropServices;

namespace PSWallpaperManagement
{
    public static class Season
    {
        public const string Spring = "Spring";
        public const string Summer = "Summer";
        public const string Fall = "Fall";
        public const string Winter = "Winter";
        public static string[] AllSeasons = { Spring, Summer, Winter, Fall };
    }

    [Cmdlet(VerbsCommon.Get, "Season")]
    public class GetSeasonCommand : PSCmdlet
    {
        private string _season { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            int month = DateTime.Now.Month;
            if (month == 3 || month == 4 || month == 5)
                _season = Season.Spring;
            else if (month == 6 || month == 7 || month == 8)
                _season = Season.Summer;
            else if (month == 9 || month == 10 || month == 11)
                _season = Season.Fall;
            else
                _season = Season.Winter;
            WriteVerbose($"Current season: {_season}");
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            WriteObject(_season);
        }
    }

    [Cmdlet(VerbsCommon.New, "WallpaperFolder")]
    public class NewWallpaperFolderCommand : PSCmdlet
    {
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            string picPath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            WriteVerbose($"Attempting to create directory at '{picPath}'");
            _ = Directory.CreateDirectory(picPath);
            foreach (string season in Season.AllSeasons)
            {
                string seasonPath = Path.Combine(picPath, season);
                WriteVerbose($"Attempting to create directory at '{seasonPath}'");
                WriteObject(Directory.CreateDirectory(seasonPath));
            }
        }
    }

    [Cmdlet(VerbsCommon.Get, "RandomWallpaper")]
    public class GetRandomWallpaperCommand : PSCmdlet
    {
        [Parameter(
            Mandatory = true,
            ValueFromPipeline = true,
            HelpMessage = "Path to a directory that contains image files."
        )]
        public string Path { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            if (Directory.Exists(Path))
            {
                string[] files = Directory.GetFiles(Path);
                int randomIndex = new Random().Next(files.Length);
                WriteObject(files[randomIndex]);
            }
        }
    }

    [Cmdlet(VerbsCommon.Set, "Wallpaper")]
    public class SetWallpaperCommand : PSCmdlet
    {
        // I don't know what this does but it works.
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(
            uint action, uint uParam, string vParam, uint winIni);

        private static readonly uint SPI_SETDESKWALLPAPER = 0x14;
        private static readonly uint SPIF_UPDATEINIFILE = 0x01;
        private static readonly uint SPIF_SENDWININICHANGE = 0x02;

        private void SetWallpaper(string path)
        {
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path,
                SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
        }
        // End mystery code.

        [Parameter(Mandatory = true, ValueFromPipeline = true)]
        public string Path { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            if (File.Exists(Path))
            {
                WriteVerbose($"Setting wallpaper to '{Path}'");
                SetWallpaper(Path);
            }
            else
            {
                WriteVerbose($"File does not exist at '{Path}'.");
            }
        }
    }
}
