/*
 * This module is build on top of on J.Bradshaw's vmcController
 * Allows some commands to not be available (e.g. the Vista EPG commands on Windows 7)
 *    In theory this should never be used...
 * 
 * Copyright (c) 2009 Anthony Jones
 * 
 * Portions copyright (c) 2007 Jonathan Bradshaw
 * 
 * This software code module is provided 'as-is', without any express or implied warranty. 
 * In no event will the authors be held liable for any damages arising from the use of this software.
 * Permission is granted to anyone to use this software for any purpose, including commercial 
 * applications, and to alter it and redistribute it freely.
 * 
 * History:
 * 2009-09-14 Created by Anthony Jones
 * 
 */
using System;

namespace VmcController.AddIn.Commands
{
	/// <summary>
	/// Summary description for MsgBox command.
	/// </summary>
	public class UnavailableCmd : ICommand
	{
        string err_text = "";

        #region ICommand Members

        public UnavailableCmd(string e)
        {
            err_text = e;
        }

        /// <summary>
        /// Shows the syntax.
        /// </summary>
        /// <returns></returns>
        public string ShowSyntax()
        {
            return "Command Unavailable: " + err_text;
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
            opResult.StatusCode = OpStatusCode.BadRequest;
            opResult.StatusText = "This command unavailable due to exception at load: " + err_text;
            return opResult;
        }

        #endregion
    }
}
