using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
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
    }

    public class OpResult
    {
        private OpResultObject m_resultObject;        
        private OpStatusCode m_statusCode = OpStatusCode.BadRequest;
        private string m_statusText = "";
        private StringBuilder m_content;
        private string m_serializedText;
        private byte[] imageData;
        private int result_count = 0;
        private bool help_format = false;


        public class DefaultOpResultObject : OpResultObject
        {
            public string message = "";
        }

        public OpResult() { }

        public OpResult(OpStatusCode statusCode)
        {
            m_statusCode = statusCode;
        }

        public OpResult(string serializedText)
        {
            m_serializedText = serializedText;
        }

        public bool isHelpFormat()
        {
            return help_format;
        }

        public OpStatusCode StatusCode
        {
            get { return m_statusCode; }
            set { m_statusCode = value; }
        }

        public string StatusText
        {
            get { return m_statusText; }
            set { m_statusText = value; }
        }

        public OpResultObject ContentObject
        {
            set { m_resultObject = value; }
        }

        public int ResultCount
        {
            set { result_count = value; }
            get { return result_count; }
        }

        public int Length
        {
            get 
            {
                if (help_format)
                {
                    return m_content.Length;
                }
                else
                {
                    return ToString().Length; 
                }                
            }
        }

        public void AppendFormat(string format, params object[] args)
        {
            if (m_content == null)
            {
                m_content = new StringBuilder();
                help_format = true;
            }
            m_content.AppendFormat(format, args);
            m_content.AppendLine();
        }


        public byte[] ImageData
        {
            set { imageData = value; }
            get { return imageData; }
        }

        public override string ToString()
        {
            if (!help_format)
            {
                if (m_serializedText == null)
                {
                    if (m_resultObject == null)
                    {
                        if (m_content != null)
                        {
                            m_resultObject = new DefaultOpResultObject();
                            ((DefaultOpResultObject)m_resultObject).message = m_content.ToString();
                        }
                        else
                        {
                            m_resultObject = new OpResultObject();
                        }
                    }
                    m_resultObject.status_code = m_statusCode;
                    if (m_statusText.Length > 0) m_resultObject.status_message = m_statusText;
                    if (result_count != 0) m_resultObject.result_count = result_count;
                    m_serializedText = JsonConvert.SerializeObject(m_resultObject, Formatting.Indented, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                }
                return m_serializedText;
            }
            else
            {
                return m_content.ToString();
            }
        }        
    }
}