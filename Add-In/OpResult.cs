using System;
using System.Collections.Generic;
using System.Net;
using System.Text;

namespace VmcController.AddIn
{
    public enum OpStatusCode
    {
        Ok = 200,
        Success = 204,
        OkImage = 208,
        BadRequest = 400,
        Exception = 500,
        Json = 800
    }

    public class OpResult
    {
        private OpStatusCode m_statusCode = OpStatusCode.BadRequest;
        private string m_statusText = string.Empty;
        private StringBuilder m_content = new StringBuilder();

        public OpResult() { }

        public OpResult(OpStatusCode statusCode)
        {
            m_statusCode = statusCode;
        }

        public string StatusText
        {
            get { return (m_statusText.Length > 0) ? m_statusText : m_statusCode.ToString(); }
            set { m_statusText = value; }
        }

        public OpStatusCode StatusCode
        {
            get { return m_statusCode; }
            set { m_statusCode = value; }
        }

        public string ContentText
        {
            set { m_content.Remove(0, m_content.Length); m_content.Append(value); }
        }

        public int Length
        {
            get { return m_content.Length; }
        }

        public void AppendFormat(string format, params object[] args)
        {
            m_content.AppendFormat(format, args);
            m_content.AppendLine();
        }

        public override string ToString()
        {
            return m_content.ToString();
        }

    }
}
