﻿namespace AllocatingStuff
{
    #region

    using System;
    using System.Windows.Forms;

    #endregion

    /// <summary>
    /// The program.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}