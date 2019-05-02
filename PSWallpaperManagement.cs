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

        public static string GetCurrentSeason()
        {
            int month = DateTime.Now.Month;
            if (month == 3 || month == 4 || month == 5)
                return Spring;
            else if (month == 6 || month == 7 || month == 8)
                return Summer;
            else if (month == 9 || month == 10 || month == 11)
                return Fall;
            else
                return Winter;
        }
    }

    [Cmdlet(VerbsCommon.Get, "Season")]
    public class GetSeasonCommand : PSCmdlet
    {
        private string _season;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            _season = Season.GetCurrentSeason();
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
        private string _picPath;

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            _picPath = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            WriteVerbose($"Attempting to create directory at '{_picPath}'");
            _ = Directory.CreateDirectory(_picPath);
            foreach (string season in Season.AllSeasons)
            {
                string seasonPath = Path.Combine(_picPath, season);
                WriteVerbose($"Attempting to create directory at '{seasonPath}'");
                WriteObject(Directory.CreateDirectory(seasonPath));
            }
        }
    }

    [Cmdlet(VerbsCommon.Get, "RandomWallpaper")]
    public class GetRandomWallpaperCommand : PSCmdlet
    {
        private Random _rand;

        [Parameter(
            Mandatory = true,
            ValueFromPipeline = true,
            HelpMessage = "Path to a directory that contains image files."
        )]
        public string Path { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            _rand = new Random();
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            if (Directory.Exists(Path))
            {
                string[] files = Directory.GetFiles(Path);
                int randomIndex = _rand.Next(files.Length);
                WriteObject(files[randomIndex]);
            }
        }
    }

    [Cmdlet(VerbsCommon.Set, "Wallpaper")]
    public class SetWallpaperCommand : PSCmdlet
    {
        /// <summary>
        /// Retrieves or sets the value of one of the system-wide parameters.
        /// This function can also update the user profile while setting a parameter.
        /// </summary>
        /// <param name="uiAction">
        /// The system-wide parameter to be retrieved or set.
        /// </param>
        /// <param name="uiParam">
        /// A parameter whose usage and format depends on the system parameter being queried or set.
        /// </param>
        /// <param name="pvParam">
        /// A parameter whose usage and format depends on the system parameter being queried or set.
        /// </param>
        /// <param name="fWinIni">
        /// If a system parameter is being set, specifies whether the user profile is to be updated,
        /// and if so, whether the WM_SETTINGCHANGE message is to be broadcast to all top-level 
        /// windows to notify them of the change.
        /// </param>
        /// <returns></returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(
            uint uiAction, uint uiParam, string pvParam, uint fWinIni);

        private static readonly uint SPI_SETDESKWALLPAPER = 0x14;
        private static readonly uint SPIF_UPDATEINIFILE = 0x01;
        private static readonly uint SPIF_SENDWININICHANGE = 0x02;

        private void SetWallpaper(string path)
        {
            _ = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path,
                    SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
        }

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
