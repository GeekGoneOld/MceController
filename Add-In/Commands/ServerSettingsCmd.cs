using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.Reflection;
/*
 * Copyright (c) 2013 Skip Mercier
 * 
 * This software is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely, subject to the following restrictions:
 * 
 * 1. The origin of this software must not be misrepresented; you must not claim that you wrote 
 *    the original software. If you use this software in a product, an acknowledgment in the 
 *    product documentation would be appreciated but is not required.
 * 2. Altered source versions must be plainly marked as such, and must not be misrepresented as
 *    being the original software.
 * 3. This notice may not be removed or altered from any source distribution.
 * 
 */
using System.Text;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for FullScreen command.
	/// </summary>
	public class ServerSettingsCmd : ICommand
	{
        private ServerSettings m_settings;

        #region ICommand Members

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<set-build-hour:hour|set-send-buffer:size> Without params, gets the server settings: version, is_building, cache_build_hour, send_buffer_size, is_cache_outdated");          
            sb.AppendLine("With params, hour is >= 0 and < 24 and size is an integer greater than " + HttpSocketServer.MIN_SEND_BUFFER_SIZE);
            return sb.ToString();
        }        

        public void set(ServerSettings settings)
        {
            m_settings = settings;
        }
        
        /// <summary>
        /// Executes the specified param.
        /// </summary>
        /// <param name="param">The param.</param>
        /// <param name="result">The result.</param>
        /// <returns></returns>
        public OpResult Execute(string param)
        {
            OpResult opResult = new OpResult();
            opResult.StatusCode = OpStatusCode.Success;
            opResult.ContentObject = m_settings;
            return opResult;
        }        

        #endregion
    }
}
