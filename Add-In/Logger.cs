using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;


namespace VmcController.AddIn
{
    public class Logger : IDisposable
    {
        protected string _filename;
        protected StreamWriter _writer;
        protected bool _disposed;

        /// <summary>
        /// Gets or sets if the file is closed between each call to Write()
        /// </summary>
        public bool KeepOpen { get; set; }

        /// <summary>
        /// Gets or sets if logging is active or suppressed
        /// </summary>
        public bool IsLogging { get; set; }

        /// <summary>
        /// Constructs a new LogFile object. Data will be appended to
        /// any existing data. File will be closed between writes.
        /// </summary>
        /// <param name="filename">Name of log file to write.</param>
        public Logger(string filename)
            : this(filename, false)
        {
        }

        /// <summary>
        /// Constructs a new LogFile object.
        /// </summary>
        /// <param name="filename">Name of log file to write.</param>
        /// <param name="append">If true, data is written to the end of any
        /// existing data; otherwise, existing data is overwritten.</param>
        /// <param name="keepOpen">If true, performance is improved by
        /// keeping the file open between writes; otherwise, the file
        /// is opened and closed for each write.</param>
        public Logger(string source, bool keepOpen)
        {
            _filename = AddInModule.DATA_DIR + "\\" + source + ".log";
            KeepOpen = keepOpen;
            IsLogging = true;

            _writer = null;
            _disposed = false;
        }

        /// <summary>
        /// Closes the current log file
        /// </summary>
        public void Close()
        {
            if (_writer != null)
            {
                _writer.Dispose();
                _writer = null;
            }
        }

        /// <summary>
        /// Writes a formatted string to the current log file
        /// </summary>
        /// <param name="fmt"></param>
        /// <param name="args"></param>
        public void Write(string fmt, params object[] args)
        {
            Write(String.Format(fmt, args));
        }

        /// <summary>
        /// Writes a string to the current log file
        /// </summary>
        /// <param name="s"></param>
        public void Write(string s)
        {
            if (IsLogging)
            {
                // Establish file stream if needed
                if (_writer == null)
                    _writer = new StreamWriter(_filename, true);

                // Write string with date/time stamp
                _writer.WriteLine(String.Format("{0:d} {0:T} : {1}", DateTime.Now, s));

                // Close file if not keeping open
                if (!KeepOpen)
                    Close();
            }
        }

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // Need to dispose managed resources if being called manually
                if (disposing)
                    Close();
                _disposed = true;
            }
        }

        #endregion
    }
}
